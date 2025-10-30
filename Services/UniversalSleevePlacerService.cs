using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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

        // ⚠️ QUICK WIN: Pre-cached family symbols (load once, reuse many times)
        private static Dictionary<string, FamilySymbol> _familySymbolCache = new Dictionary<string, FamilySymbol>();

        public int PlacedCount { get; private set; }
        public int SkippedCount { get; private set; }
        public int ErrorCount { get; private set; }

        public UniversalSleevePlacerService(Document doc, OpeningConditions conditions, ISleevePlacementStrategy strategy, Dictionary<string, double> clearanceSettings = null, string filterName = null)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _conditions = conditions ?? new OpeningConditions();
            _strategy = strategy ?? throw new ArgumentNullException(nameof(strategy));
            _clearanceSettings = clearanceSettings ?? new Dictionary<string, double>();
            _filterName = filterName;
            
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
            
            DebugLogger.Info($"[UniversalSleevePlaycer] Initialized for category: {_strategy.GetCategoryName()}");
            DebugLogger.Info($"[UniversalSleevePlaycer] Received {_clearanceSettings.Count} clearance settings from UI");
            
            // Log all clearance settings for debugging
            foreach (var kvp in _clearanceSettings)
            {
                DebugLogger.Info($"[UniversalSleevePlaycer] Clearance: {kvp.Key} = {kvp.Value}mm");
            }
            
            // Add build timestamp to placement_debug.log
            try
            {
                var asm = System.Reflection.Assembly.GetExecutingAssembly();
                var ver = System.Diagnostics.FileVersionInfo.GetVersionInfo(asm.Location)?.FileVersion ?? "?";
                var ts = System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
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
                
                DebugLogger.Info($"[UniversalSleevePlacer] Looking for {mepElementIds.Count} MEP elements with IDs from clash zones");
                
                // 1. Search in active document first
                var activeElements = new FilteredElementCollector(_doc)
                    .WhereElementIsNotElementType()
                    .Where(elem => mepElementIds.Contains(elem.Id))
                    .ToList();
                
                mepElements.AddRange(activeElements);
                DebugLogger.Info($"[UniversalSleevePlacer] Found {activeElements.Count} MEP elements in active document");
                
                // 2. Search in linked documents
                var linkedDocs = _doc.Application.Documents.Cast<Document>()
                    .Where(doc => doc.IsLinked && doc.Title != _doc.Title)
                    .ToList();
                
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
                        DebugLogger.Info($"[UniversalSleevePlacer] Found {linkedElements.Count} MEP elements in linked document: {linkedDoc.Title}");
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Warning($"[UniversalSleevePlacer] Error searching MEP elements in linked document {linkedDoc.Title}: {ex.Message}");
                    }
                }
                
                DebugLogger.Info($"[UniversalSleevePlacer] Total collected {mepElements.Count} MEP elements from all documents");
                return mepElements;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UniversalSleevePlacer] Error collecting MEP elements: {ex.Message}");
                return new List<Element>();
            }
        }
        
        /// <summary>
        /// STEP 2: Place all sleeves within an active transaction
        /// WRITE phase - Transaction MUST be started by caller (Command)
        /// Matches MD pattern: Section 2, Lines 72-80
        /// </summary>
        public (int PlacedCount, int SkippedCount) PlaceAllSleevesInTransaction(List<ClashZone> clashZones)
        {
            // ⏱️ TIMING: Start overall placement timer
            var overallTimer = System.Diagnostics.Stopwatch.StartNew();
            var detailedTimingLog = new System.Text.StringBuilder();
            
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
            
            // ✅ PERFORMANCE OPTIMIZATION: Cache GlobalFlagManager per category (not per clash zone)
            var globalManagersByCategory = new Dictionary<string, GlobalFlagManager>();
            
            // ✅ PERFORMANCE OPTIMIZATION: Batch file logging - collect logs and write once
            var batchLogs = new System.Text.StringBuilder();
            
            // 🔥 CRITICAL DEBUG: Log flag status from the clash zones passed to this method
            int clusterResolvedCount = clashZones.Count(cz => cz.IsClusterResolved);
            int individualResolvedCount = clashZones.Count(cz => cz.IsResolved);
            batchLogs.AppendLine($"[{DateTime.Now:HH:mm:ss}] 🔥 UNIVERSAL SLEEVE PLACER RECEIVED: {clashZones.Count} clash zones");
            batchLogs.AppendLine($"[{DateTime.Now:HH:mm:ss}] 📊 FLAGS RECEIVED: IsClusterResolved=True: {clusterResolvedCount}, IsResolved=True: {individualResolvedCount}");
            
            // 🛡️ FAIL-SAFE: Check document state before starting
            if (!_doc.IsModifiable)
            {
                DebugLogger.Error("[UniversalSleevePlacer] Document is read-only or workshared and not checked out");
                throw new InvalidOperationException("Document is not modifiable. Please check out the file or ensure it's not read-only.");
            }
            
            if (clashZones == null || clashZones.Count == 0)
            {
                DebugLogger.Warning($"[UniversalSleevePlacer] No clash zones provided");
                return (0, 0);
            }
            
            // DEBUG: Log all incoming ClashZone objects to trace XML loading
            DebugLogger.Log($"[XML-DEBUG] PlaceAllSleevesInTransaction called with {clashZones.Count} clash zones:");
            foreach (var cz in clashZones.Take(5)) // Log first 5 to avoid spam
            {
                DebugLogger.Log($"[XML-DEBUG] ClashZone {cz.Id}: MepElementCategory='{cz.MepElementCategory}', StructuralElementType='{cz.StructuralElementType}', MepElementId={cz.MepElementIdValue}, StructuralElementId={cz.StructuralElementIdValue}");
            }
            
            // ⚠️ DON'T reset resolved flags during placement!
            // Flags are managed by refresh - it checks if sleeves exist and resets flags if deleted
            // If we reset here, we'll place duplicate sleeves for existing ones
            DebugLogger.Info($"[UniversalSleevePlacer] Processing {clashZones.Count} clash zones (zero linked file access), trusting IsResolved flags from refresh");
            
            // ✅ PRIORITY SORTING: Process Duct Accessories (Dampers) BEFORE Ducts
            var sortedClashZones = clashZones
                .OrderBy(cz => GetCategoryPriority(cz.MepElementCategory))
                .ThenBy(cz => cz.Id)
                .ToList();
            
            DebugLogger.Info($"[UniversalSleevePlacer] Sorted {sortedClashZones.Count} clash zones by category priority");

            // Log first 10 sorted clash zones to verify order
            DebugLogger.Info("[SORT-VERIFY] First 10 clash zones after sorting:");
            foreach (var cz in sortedClashZones.Take(10))
            {
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
                foreach (var clashZone in sortedClashZones)
                {
                    // ⏱️ TIMING: Start per-sleeve timer (only for actual placement operations)
                    var sleeveTimer = System.Diagnostics.Stopwatch.StartNew();
                    var sleeveLog = new System.Text.StringBuilder();
                    
                    try
                    {
                        DebugLogger.Info($"[UniversalSleevePlacer] Processing ClashZone {clashZone.Id}: MEP={clashZone.MepElementId.IntegerValue}, Structural={clashZone.StructuralElementId.IntegerValue}");
                        
                        // ✅ PERFORMANCE OPTIMIZATION: Batch file logging instead of individual writes
                        // Only log to batch - will write once at end (or every 50 clash zones)
                        if (PlacedCount + SkippedCount < 50) // Only log first 50 for debugging
                        {
                            batchLogs.AppendLine($"[{DateTime.Now:HH:mm:ss}] [SLEEVE-PLACER] ClashZone {clashZone.Id}: MEP={clashZone.MepElementId.IntegerValue}, Structural={clashZone.StructuralElementId.IntegerValue}");
                            batchLogs.AppendLine($"[{DateTime.Now:HH:mm:ss}] [SLEEVE-PLACER] FLAGS: IsResolved={clashZone.IsResolved}, IsClusterResolved={clashZone.IsClusterResolved}");
                        }
                        
                        // STEP 0: Check if cluster sleeve actually exists (prevent individual sleeves over cluster sleeves)
                        if (clashZone.IsClusterResolved || clashZone.ClusterSleeveInstanceId > 0)
                        {
                            DebugLogger.Info($"[UniversalSleevePlacer] SKIP: ClashZone {clashZone.Id} has cluster sleeve {clashZone.ClusterSleeveInstanceId} - preventing individual sleeve placement");
                            if (PlacedCount + SkippedCount < 50) batchLogs.AppendLine($"✓ SKIP ClashZone {clashZone.Id}: has cluster sleeve {clashZone.ClusterSleeveInstanceId}");
                            SkippedCount++;
                            sleeveTimer.Stop();
                            continue;
                        }
                        
                        // STEP 1: Check individual sleeve flag first (fastest check)
                        if (clashZone.IsResolved)
                        {
                            // ✅ CRITICAL: Verify sleeve actually exists in Revit before trusting XML flags
                            if (clashZone.SleeveInstanceId > 0)
                            {
                                try
                                {
                                    var sleeveElement = _doc.GetElement(new ElementId(clashZone.SleeveInstanceId));
                                    if (sleeveElement == null)
                                    {
                                        // Sleeve was deleted - reset flags and continue with placement
                                        DebugLogger.Info($"[UniversalSleevePlacer] RESET: Sleeve {clashZone.SleeveInstanceId} was deleted - resetting flags for ClashZone {clashZone.Id}");
                                        if (PlacedCount + SkippedCount < 50) batchLogs.AppendLine($"🔄 RESET ClashZone {clashZone.Id}: Sleeve {clashZone.SleeveInstanceId} was deleted - resetting flags");
                                        clashZone.IsResolved = false;
                                        clashZone.SleeveInstanceId = -1;
                                        clashZone.SleeveFamilyName = string.Empty;
                                        // Continue to placement (don't skip)
                                    }
                                    else
                                    {
                                        // Sleeve exists - check both flags according to flag management logic
                            if (clashZone.IsClusterResolved)
                            {
                                // Both flags TRUE: Skip processing (avoid clash zone)
                                DebugLogger.Info($"[DEBUG] ✓ SKIP ClashZone {clashZone.Id}: Both IsResolved=True AND IsClusterResolved=True - avoid clash zone\n");
                                SkippedCount++;
                                continue;
                            }
                            else
                            {
                                // Only individual flag TRUE: Skip individual sleeve placement
                                DebugLogger.Info($"[DEBUG] ✓ SKIP ClashZone {clashZone.Id}: IsResolved=True (individual sleeve already placed)\n");
                                SkippedCount++;
                                continue;
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    DebugLogger.Error($"[UniversalSleevePlacer] Error checking sleeve existence: {ex.Message}");
                                    // On error, reset flags and continue with placement
                                    clashZone.IsResolved = false;
                                    clashZone.SleeveInstanceId = -1;
                                    clashZone.SleeveFamilyName = string.Empty;
                                }
                            }
                            else
                            {
                                // Invalid SleeveInstanceId - reset flags and continue with placement
                                DebugLogger.Info($"[UniversalSleevePlacer] RESET: Invalid SleeveInstanceId {clashZone.SleeveInstanceId} - resetting flags for ClashZone {clashZone.Id}");
                                if (PlacedCount + SkippedCount < 50) batchLogs.AppendLine($"🔄 RESET ClashZone {clashZone.Id}: Invalid SleeveInstanceId {clashZone.SleeveInstanceId} - resetting flags");
                                clashZone.IsResolved = false;
                                clashZone.SleeveInstanceId = -1;
                                clashZone.SleeveFamilyName = string.Empty;
                                // Continue to placement (don't skip)
                            }
                        }
                        
                        // STEP 2: Check cluster sleeve flag
                        if (clashZone.IsClusterResolved)
                        {
                            // Only cluster flag TRUE: Skip both individual and cluster placement
                            if (PlacedCount + SkippedCount < 50) batchLogs.AppendLine($"✓ SKIP ClashZone {clashZone.Id}: IsClusterResolved=True (cluster sleeve handles it)");
                            SkippedCount++;
                            continue;
                        }
                        
                        // STEP 3: Both flags FALSE - proceed with individual sleeve placement
                        
                        // STEP 4: Check if individual sleeve already exists in Revit (additional safety check)
                        if (clashZone.SleeveInstanceId > 0)
                        {
                            // Check if individual sleeve still exists in Revit
                            var individualSleeveId = new ElementId(clashZone.SleeveInstanceId);
                            var individualSleeve = _doc.GetElement(individualSleeveId);
                            
                            if (individualSleeve != null)
                            {
                                // Individual sleeve exists - skip
                                DebugLogger.Info($"[UniversalSleevePlacer] SKIP: ClashZone {clashZone.Id} already has individual sleeve {clashZone.SleeveInstanceId}");
                                SkippedCount++;
                                continue;
                            }
                            else
                            {
                                // ✅ CRITICAL FIX: Only reset individual flag if cluster flag is also false
                                if (!clashZone.IsClusterResolved)
                                {
                                    // No cluster sleeve - reset individual flag and place new one
                                    DebugLogger.Info($"[UniversalSleevePlacer] Individual sleeve {clashZone.SleeveInstanceId} was deleted and no cluster sleeve - resetting flag and placing new sleeve");
                                    clashZone.IsResolved = false;
                                    clashZone.SleeveInstanceId = -1;
                                    // Continue to place individual sleeve below
                                }
                                else
                                {
                                    // Cluster sleeve exists - keep individual flag true and skip
                                    DebugLogger.Info($"[UniversalSleevePlacer] SKIP: ClashZone {clashZone.Id} has cluster sleeve - keeping individual flag true");
                                    DebugLogger.Info($"[DEBUG] ✓ SKIP ClashZone {clashZone.Id}: Individual sleeve missing but cluster sleeve exists - keeping IsResolved=true\n");
                                    SkippedCount++;
                                    continue;
                                }
                            }
                        }
                        
                        // STEP 3: If we reach here, place individual sleeve (fresh or replacement)
                        
                        // ✅ CRITICAL: Check Global XML before placement (prevents cross-filter duplicates)
                        // ✅ PERFORMANCE OPTIMIZATION: Cache GlobalFlagManager per category instead of creating new one per clash zone
                        try
                        {
                            var categoryName = clashZone.MepElementCategory;
                            if (!globalManagersByCategory.TryGetValue(categoryName, out var globalManager))
                            {
                                // ✅ MEMORY OPTIMIZATION: Use static singleton to avoid reloading XML (works across all service instances)
                                globalManager = GlobalFlagManager.GetOrCreate(categoryName);
                                globalManagersByCategory[categoryName] = globalManager;
                            }
                            var sleeveState = globalManager.CheckSleeveExistence(_doc, clashZone.MepElementId, clashZone.StructuralElementId);
                            
                            if (sleeveState.ExistsInGlobal)
                            {
                                if (sleeveState.HasClusterSleeve)
                                {
                                    // Cluster sleeve exists in Global XML and Revit - skip individual placement
                                    DebugLogger.Info($"[GLOBAL-XML] SKIP: ClashZone {clashZone.Id} has cluster sleeve {sleeveState.Placement.ClusterSleeveId} in Global XML (placed from another filter)");
                                    if (PlacedCount + SkippedCount < 50) batchLogs.AppendLine($"✓ SKIP ClashZone {clashZone.Id}: Cluster sleeve {sleeveState.Placement.ClusterSleeveId} exists in Global XML");
                                    
                                    // Update clash zone flags to match Global XML state
                                    clashZone.IsClusterResolved = true;
                                    clashZone.ClusterSleeveInstanceId = sleeveState.Placement.ClusterSleeveId;
                                    clashZone.IsResolved = false;
                                    clashZone.SleeveInstanceId = -1;
                                    
                                    SkippedCount++;
                                    continue;
                                }
                                else if (sleeveState.HasIndividualSleeve)
                                {
                                    // Individual sleeve exists in Global XML and Revit - skip placement
                                    DebugLogger.Info($"[GLOBAL-XML] SKIP: ClashZone {clashZone.Id} has individual sleeve {sleeveState.Placement.IndividualSleeveId} in Global XML (placed from another filter)");
                                    if (PlacedCount + SkippedCount < 50) batchLogs.AppendLine($"✓ SKIP ClashZone {clashZone.Id}: Individual sleeve {sleeveState.Placement.IndividualSleeveId} exists in Global XML");
                                    
                                    // Update clash zone flags to match Global XML state
                                    clashZone.IsResolved = true;
                                    clashZone.SleeveInstanceId = sleeveState.Placement.IndividualSleeveId;
                                    
                                    SkippedCount++;
                                    continue;
                                }
                                else
                                {
                                    // Entry exists in Global XML but sleeve was deleted from Revit - reset Global XML entry and proceed
                                    DebugLogger.Info($"[GLOBAL-XML] RESET: ClashZone {clashZone.Id} has Global XML entry but sleeve was deleted - removing Global XML entry and proceeding with placement");
                                    globalManager.RemovePlacement(clashZone.MepElementId, clashZone.StructuralElementId);
                                    if (PlacedCount + SkippedCount < 50) batchLogs.AppendLine($"🔄 RESET ClashZone {clashZone.Id}: Global XML entry existed but sleeve deleted - removed entry, proceeding with placement");
                                }
                            }
                            // If no Global XML entry exists, proceed with placement (normal case)
                        }
                        catch (Exception globalEx)
                        {
                            DebugLogger.Warning($"[GLOBAL-XML] Error checking Global XML for ClashZone {clashZone.Id}: {globalEx.Message} - proceeding with placement");
                            // Continue with placement on error (fail-safe)
                        }
                        
                        // ✅ PERFORMANCE OPTIMIZATION: Removed excessive file logging - use DebugLogger only
                        // Validate category match
                        if (!string.IsNullOrEmpty(clashZone.MepElementCategory) && 
                            clashZone.MepElementCategory != _strategy.GetCategoryName())
                        {
                            DebugLogger.Warning($"[UniversalSleevePlacer] SKIP: ClashZone {clashZone.Id} category '{clashZone.MepElementCategory}' doesn't match '{_strategy.GetCategoryName()}'");
                            SkippedCount++;
                            sleeveTimer.Stop();
                            continue;
                        }
                        
                        
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
                        
						// ⚠️ SPECIAL HANDLING & SIZING ORDER:
						// 1) Pipes (host-agnostic), 2) Dampers, 3) Cable trays, 4) Ducts/default
                        XYZ placementOffset = XYZ.Zero;
                        double finalWidth = 0.0, finalHeight = 0.0, finalDiameter = 0.0;
                        
                        // ⏱️ TIMING: Clearance calculation
                        var clearanceTimer = System.Diagnostics.Stopwatch.StartNew();
                        // 🛡️ ARCHITECTURE FIX: Use CONDITIONS service for ALL clearance types
                        // This ensures consistent architecture: CONDITIONS XML → UniversalSleevePlacerService
                        // Raw dimensions from ClashZone + Clearance from CONDITIONS = Final dimensions
                        
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
                            
                            DebugLogger.Info($"[UniversalSleevePlacer] DAMPER: Raw={UnitUtils.ConvertFromInternalUnits(rawWidth, UnitTypeId.Millimeters):F1}x{UnitUtils.ConvertFromInternalUnits(rawHeight, UnitTypeId.Millimeters):F1}mm → Final={UnitUtils.ConvertFromInternalUnits(finalWidth, UnitTypeId.Millimeters):F1}x{UnitUtils.ConvertFromInternalUnits(finalHeight, UnitTypeId.Millimeters):F1}mm");
                        }
                        
                        // ✅ DEBUG: Log strategy information
                        DebugLogger.Info($"[UniversalSleevePlacer] Strategy Type: {_strategy.GetType().Name}");
                        DebugLogger.Info($"[UniversalSleevePlacer] Strategy Category: {_strategy.GetCategoryName()}");
                        
                        if (_strategy is DuctPlacementStrategy ductStrategy)
                        {
                            // ✅ Ducts: Raw dimensions + CONDITIONS clearance via strategy
                            var rawWidth = clashZone.MepElementWidth;
                            var rawHeight = clashZone.MepElementHeight;
                            
                            DebugLogger.Info($"[UniversalSleevePlacer] DUCT CALCULATION START: Raw dimensions {UnitUtils.ConvertFromInternalUnits(rawWidth, UnitTypeId.Millimeters):F1}x{UnitUtils.ConvertFromInternalUnits(rawHeight, UnitTypeId.Millimeters):F1}mm");
                            
                            var clearance = GetClearanceFromConditions("Ducts", mepSize);
                            finalWidth = rawWidth + (2 * clearance);
                            finalHeight = rawHeight + (2 * clearance);
                            finalDiameter = finalWidth; // For round elements
                            
                            DebugLogger.Info($"[UniversalSleevePlacer] DUCT: Raw={UnitUtils.ConvertFromInternalUnits(rawWidth, UnitTypeId.Millimeters):F1}x{UnitUtils.ConvertFromInternalUnits(rawHeight, UnitTypeId.Millimeters):F1}mm + Clearance={UnitUtils.ConvertFromInternalUnits(clearance, UnitTypeId.Millimeters):F1}mm = Final={UnitUtils.ConvertFromInternalUnits(finalWidth, UnitTypeId.Millimeters):F1}x{UnitUtils.ConvertFromInternalUnits(finalHeight, UnitTypeId.Millimeters):F1}mm");
                        }
                        else if (_strategy is CableTrayPlacementStrategy cableTrayStrategy)
                        {
                            // ✅ PERFORMANCE OPTIMIZATION: Removed excessive file logging

                            // ✅ Cable trays: Raw dimensions + UI/XML clearance via strategy
                            var rawWidth = clashZone.MepElementWidth;
                            var rawHeight = clashZone.MepElementHeight;

                            DebugLogger.Info($"[UniversalSleevePlacer] CABLE TRAY STRATEGY: Raw={UnitUtils.ConvertFromInternalUnits(rawWidth, UnitTypeId.Millimeters):F1}x{UnitUtils.ConvertFromInternalUnits(rawHeight, UnitTypeId.Millimeters):F1}mm");
                            DebugLogger.Info($"[UniversalSleevePlacer] CABLE TRAY STRATEGY: UI Clearance Settings Count={_clearanceSettings.Count}");

                            // Get offset and final dimensions from strategy (uses UI settings first, then CONDITIONS)
                            var adj2 = cableTrayStrategy.GetCableTrayPlacementAdjustment(clashZone, _conditions, _clearanceSettings);
                            placementOffset = adj2.offsetVector;
                            finalWidth = adj2.finalWidth;
                            finalHeight = adj2.finalHeight;

                            DebugLogger.Info($"[UniversalSleevePlacer] CABLE TRAY STRATEGY: Final={UnitUtils.ConvertFromInternalUnits(finalWidth, UnitTypeId.Millimeters):F1}x{UnitUtils.ConvertFromInternalUnits(finalHeight, UnitTypeId.Millimeters):F1}mm");
                            finalDiameter = finalWidth; // Not used for cable trays (rectangular only)

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
                            
                            DebugLogger.Warning($"[UniversalSleevePlacer] ⚠️ NO STRATEGY MATCHED for ClashZone {clashZone.Id} - Using fallback dimensions");
                        }
                        clearanceTimer.Stop();
                        totalClearanceTime += clearanceTimer.Elapsed;
                        sleeveLog.AppendLine($"  Clearance calc: {clearanceTimer.ElapsedMilliseconds}ms");

                        // ⚠️ REMOVED: Old width/height swapping logic that was causing double-swapping
                        // The new logic later in the method (lines 660-665) handles this correctly
                        // by ensuring the longer dimension becomes width, not just swapping blindly
                        
                        // ⏱️ TIMING: Family selection and loading
                        var familyTimer = System.Diagnostics.Stopwatch.StartNew();
                        // Select universal family
                        var (familyName, typeName, isCircular) = SelectUniversalFamily(clashZone, mepSize);
                        var familySymbol = LoadFamilySymbol(familyName);
                        familyTimer.Stop();
                        totalFamilyLoadTime += familyTimer.Elapsed;
                        sleeveLog.AppendLine($"  Family load: {familyTimer.ElapsedMilliseconds}ms");
                        
                        if (familySymbol == null)
                        {
                            DebugLogger.Error($"[UniversalSleevePlacer] Family '{familyName}' not found");
                            ErrorCount++;
                            continue;
                        }
                        
                        // Activate symbol if needed
                        if (!familySymbol.IsActive)
                        {
                            familySymbol.Activate();
                        }
                        
                        // Use the exact clash placement point (no substitution) and apply strategy offset
                        var placementPointChosen = clashZone.SleevePlacementPoint;
                        
                        DebugLogger.Info($"[UniversalSleevePlacer] Using exact intersection point {placementPointChosen} for {clashZone.MepElementCategory} on {clashZone.StructuralElementType}");

                        XYZ adjustedPlacementPoint = placementPointChosen + placementOffset;
                        DebugLogger.Info($"[UniversalSleevePlacer] Using placement point {placementPointChosen} → adjusted {adjustedPlacementPoint}");
                        
                        if (placementOffset.GetLength() > 0.001)
                        {
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
                                DebugLogger.Warning($"[UniversalSleevePlacer] No level found BELOW placement point Z={placementZ:F3} for {clashZone.StructuralElementType}");
                                SkippedCount++;
                                sleeveTimer.Stop();
                                continue;
                            }
                            
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
                            DebugLogger.Warning($"[UniversalSleevePlacer] No level found for placement point");
                            SkippedCount++;
                                sleeveTimer.Stop();
                            continue;
                        }
                        
                            DebugLogger.Info($"[UniversalSleevePlacer] Found nearest level '{nearestLevel.Name}' (Elevation={nearestLevel.Elevation:F3}) for Floor at Z={adjustedPlacementPoint.Z:F3}");
                        }
                        levelTimer.Stop();
                        totalLevelFindTime += levelTimer.Elapsed;
                        sleeveLog.AppendLine($"  Level find: {levelTimer.ElapsedMilliseconds}ms");
                        
                        // ⏱️ TIMING: Sleeve creation
                        var createTimer = System.Diagnostics.Stopwatch.StartNew();
                        // Place sleeve instance (NO HOST PARAMETER - workplane-based families)
                        // ✅ Works with linked structural elements because no host reference needed
                        var sleeveInstance = _doc.Create.NewFamilyInstance(
                            adjustedPlacementPoint,
                            familySymbol,
                            nearestLevel,
                            StructuralType.NonStructural);
                        createTimer.Stop();
                        totalSleeveCreateTime += createTimer.Elapsed;
                        sleeveLog.AppendLine($"  Sleeve create: {createTimer.ElapsedMilliseconds}ms");
                        
                        if (sleeveInstance == null)
                        {
                            DebugLogger.Error($"[UniversalSleevePlacer] Failed to create sleeve instance");
                            ErrorCount++;
                            sleeveTimer.Stop();
                            continue;
                        }
                        
                        // ⏱️ TIMING: Parameter setting
                        var parameterTimer = System.Diagnostics.Stopwatch.StartNew();
                        // Set parameters
                        SetSleeveParameters(sleeveInstance, mepSize, finalWidth, finalHeight, finalDiameter, clashZone, isCircular);
                        
                        // Set sleeve metadata for fast parameter transfer
                        SetSleeveMetadata(sleeveInstance, clashZone.MepElementCategory);
                        
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
                                    DebugLogger.Info($"[UniversalSleevePlacer] Moved sleeve {sleeveInstance.Id} to exact placement point {adjustedPlacementPoint} (from {currentPt})");
                                }
                                else
                                {
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
                                    DebugLogger.Error($"[UniversalSleevePlacer] ❌ INVALID DIMENSIONS DETECTED for sleeve {sleeveInstance.Id.IntegerValue}!");
                                    ErrorCount++;
                                    sleeveTimer.Stop();
                                    validationTimer.Stop();
                                    continue;
                                }
                                
                                clashZone.SleeveWidth = finalWidth;
                                clashZone.SleeveHeight = finalHeight;
                                clashZone.SleeveDiameter = finalDiameter;
                                
                                // ✅ PERFORMANCE OPTIMIZATION: Removed Thread.Sleep(100) - Revit API calls are synchronous
                                // Get actual sleeve bounding box coordinates immediately
                                var actualBbox = sleeveInstance.get_BoundingBox(null);
                                if (actualBbox != null)
                                {
                                    clashZone.SleevePlacementPoint = new XYZ(
                                        (actualBbox.Min.X + actualBbox.Max.X) / 2,
                                        (actualBbox.Min.Y + actualBbox.Max.Y) / 2,
                                        (actualBbox.Min.Z + actualBbox.Max.Z) / 2
                                    );
                                    
                                    // ✅ CRITICAL FIX: Set the actual Revit element ID
                                    clashZone.SleeveInstanceId = sleeveInstance.Id.IntegerValue;
                                }
                                
                                // ✅ PERFORMANCE OPTIMIZATION: Batch XML updates instead of updating per sleeve
                                // XML will be updated once at the end of placement via orchestrator
                                
                                DebugLogger.Info($"[UniversalSleevePlacer] Updated ClashZone {clashZone.Id} with actual sleeve coordinates: {currentPt}, W={UnitUtils.ConvertFromInternalUnits(finalWidth, UnitTypeId.Millimeters):F1}mm, H={UnitUtils.ConvertFromInternalUnits(finalHeight, UnitTypeId.Millimeters):F1}mm");
                            }
                            else
                            {
                                DebugLogger.Error($"[UniversalSleevePlacer] ❌ Failed to get LocationPoint for sleeve {sleeveInstance.Id} - cannot update coordinates!");
                                DebugLogger.Info($"[COORD-UPDATE-FAIL] Sleeve {sleeveInstance.Id.IntegerValue}: LocationPoint is NULL ❌\n");
                            }
                        }
                        catch (Exception coordEx)
                        {
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
                        
                        // ✅ GLOBAL XML: Record placement in global XML (use cached manager)
                        try
                        {
                            var categoryName = clashZone.MepElementCategory;
                            if (!globalManagersByCategory.TryGetValue(categoryName, out var globalManager))
                            {
                                // ✅ MEMORY OPTIMIZATION: Use static singleton to avoid reloading XML (works across all service instances)
                                globalManager = GlobalFlagManager.GetOrCreate(categoryName);
                                globalManagersByCategory[categoryName] = globalManager;
                            }
                            
                            // Get filter filename from constructor parameter or use default
                            string filterName = _filterName ?? "unknown_filter.xml";
                            
                            globalManager.RecordPlacement(
                                clashZone.MepElementId,
                                clashZone.StructuralElementId,
                                sleeveInstance.Id,
                                null, // Individual sleeve, no cluster
                                filterName
                            );
                            
                            DebugLogger.Info($"[GLOBAL-XML] Recorded individual sleeve {sleeveInstance.Id.IntegerValue} for MEP={clashZone.MepElementId.IntegerValue}, Host={clashZone.StructuralElementId.IntegerValue}");
                        }
                        catch (Exception globalEx)
                        {
                            DebugLogger.Warning($"[GLOBAL-XML] Error recording placement: {globalEx.Message}");
                        }
                        
                        // ✅ PERFORMANCE OPTIMIZATION: Batch logging instead of individual file writes
                        if (PlacedCount < 50) batchLogs.AppendLine($"[SLEEVE-PLACED] ClashZone {clashZone.Id}: SleeveInstanceId = {clashZone.SleeveInstanceId}, RevitElementId = {sleeveInstance.Id.IntegerValue}");
                        
                        DebugLogger.Info($"[UniversalSleevePlacer] ✅ SLEEVE PLACED: ClashZone {clashZone.Id} → SleeveInstanceId = {clashZone.SleeveInstanceId} (Revit: {sleeveInstance.Id.IntegerValue})");
                        
                        // ✅ CRITICAL: SleevePlacementPoint is already set to actual sleeve location above
                        // DO NOT overwrite it with adjustedPlacementPoint - we need the REAL coordinates for clustering
                        
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
                        DebugLogger.Info($"[UniversalSleevePlacer] ✓ Placed {_strategy.GetCategoryName()} sleeve {sleeveInstance.Id} for ClashZone {clashZone.Id} at {adjustedPlacementPoint} ({sleeveTimer.ElapsedMilliseconds}ms)");
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Error($"[UniversalSleevePlacer] Error placing sleeve for ClashZone {clashZone.Id}: {ex.Message}");
                        ErrorCount++;
                        // Stop timer even on error
                        if (sleeveTimer.IsRunning) sleeveTimer.Stop();
                    }
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
                DebugLogger.Info($"[TIMING] Sleeve placement timing logged to file");
                
                DebugLogger.Info($"[UniversalSleevePlacer] Placement loop complete - Placed: {PlacedCount}, Skipped: {SkippedCount}, Errors: {ErrorCount}");
                DebugLogger.Info($"[TIMING] Total time: {overallTimer.ElapsedMilliseconds}ms, Avg per sleeve: {avgPlacementTime:F2}ms");
                DebugLogger.Info($"[{DateTime.Now}] [PLACEMENT_COMPLETE] Placed: {PlacedCount}, Skipped: {SkippedCount}, Errors: {ErrorCount}, Total: {overallTimer.ElapsedMilliseconds}ms, Avg: {avgPlacementTime:F2}ms\n");
                
                // CRITICAL FIX: Save XML files with updated SleeveInstanceId values
                if (PlacedCount > 0)
                {
                    DebugLogger.Info($"[{DateTime.Now}] [XML_SAVE] PlacedCount > 0, calling SaveUpdatedXmlFiles\n");
                    
                    // ⚠️ CRITICAL: Log flag states BEFORE XML save
                    DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] [XML-SAVE-BEFORE] About to save XML with {PlacedCount} placed sleeves\n");
                    
                    SaveUpdatedXmlFiles(clashZones);
                    
                    // ⚠️ REMOVED: Don't update coordinates here because clustering will delete individual sleeves
                    // Individual sleeve coordinates are saved via SaveUpdatedXmlFiles above
                    // Cluster sleeve coordinates will be saved AFTER clustering in OpeningCommandOrchestrator
                    DebugLogger.Info($"[{DateTime.Now}] [COORDINATE-FIX] Individual sleeve XML saved - coordinates will be updated after clustering\n");
                    
                    // ⚠️ CRITICAL: Log flag states AFTER XML save
                    DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] [XML-SAVE-AFTER] XML save completed for {PlacedCount} placed sleeves\n");

                    // ✅ GLOBAL FLAGS: Upsert IsResolved/IsClusterResolved into per-category global index (with instance IDs)
                    try
                    {
                        var updatesByCategory = clashZones
                            .GroupBy(cz => cz.MepElementCategory)
                            .ToDictionary(g => g.Key, g => g.Select(cz => (cz.Id, cz.IsResolved, cz.IsClusterResolved, cz.SleeveInstanceId, cz.ClusterSleeveInstanceId)));

                        foreach (var kvp in updatesByCategory)
                        {
                            var categoryName = kvp.Key;
                            var updates = kvp.Value;
                            GlobalIndexService.UpsertFlagsWithIds(_doc, categoryName, updates);
                        }
                    }
                    catch (Exception upEx)
                    {
                        DebugLogger.Warning($"[GLOBAL_INDEX] Upsert after individual placement failed: {upEx.Message}");
                    }
                }
                else
                {
                    DebugLogger.Info($"[{DateTime.Now}] [XML_SAVE] PlacedCount = 0, skipping SaveUpdatedXmlFiles\n");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UniversalSleevePlacer] Error in placement loop: {ex.Message}");
                throw;
            }
            finally
            {
                // ✅ PERFORMANCE OPTIMIZATION: Write batch logs once at the end instead of per clash zone
                if (batchLogs.Length > 0)
                {
                    try
                    {
                        System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log", batchLogs.ToString());
                        if (!DeploymentConfiguration.DeploymentMode)
                            System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\flag_state_debug.log", batchLogs.ToString());
                    }
                    catch { } // Don't fail placement if logging fails
                }
            }
            
            DebugLogger.Info($"[UniversalSleevePlacer] 🎯 PLACEMENT COMPLETED: Placed={PlacedCount}, Skipped={SkippedCount}");
            
            return (PlacedCount, SkippedCount);
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
                
                DebugLogger.Info($"[REGENERATION-DEBUG] Found {allSleeves.Count} sleeves with MEP_ElementId parameter\n");
                
                // Group sleeves by category using MEP_Category parameter
                var groupedSleeves = allSleeves.GroupBy(s => GetSleeveCategoryFromParameter(s)).ToList();
                
                foreach (var group in groupedSleeves)
                {
                    var category = group.Key;
                    var sleeves = group.ToList();
                    
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
                
                DebugLogger.Info($"[REGENERATION-SAVE] Saved {sleeveDataList.Count} sleeves to {fileName}\n");
            }
            catch (Exception ex)
            {
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
                
                DebugLogger.Info($"[UniversalSleevePlacer] Loading Duct Accessories clash zones from: {xmlFilePath}");
                
                var serializer = new System.Xml.Serialization.XmlSerializer(typeof(OpeningFilter));
                using (var reader = new StreamReader(xmlFilePath))
                {
                    var filter = (OpeningFilter)serializer.Deserialize(reader);
                    if (filter?.ClashZoneStorage?.ClashZones != null)
                    {
                        clashZones.AddRange(filter.ClashZoneStorage.ClashZones);
                        DebugLogger.Info($"[UniversalSleevePlacer] Loaded {clashZones.Count} Duct Accessories clash zones from XML");
                    }
                }
            }
            else
            {
                DebugLogger.Warning($"[UniversalSleevePlacer] No Duct Accessories XML files found matching pattern: {pattern}");
            }
            
            return clashZones;
        }
        catch (Exception ex)
        {
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
                DebugLogger.Info($"[UniversalSleevePlacer] PIPE OPENING TYPE DEBUG: Host={hostType}, Category={clashZone.MepElementCategory}, UI_Preference={pipeType}");
                
                // Use PipePlacementStrategy to resolve opening type with global rules
                if (_strategy is PipePlacementStrategy pipeStrategy)
                {
                    var resolvedType = pipeStrategy.GetResolvedOpeningType(mepSize, pipeType, hostType);
                    isCircular = string.Equals(resolvedType, "Circular", StringComparison.OrdinalIgnoreCase);
                    DebugLogger.Info($"[UniversalSleevePlacer] PIPE opening type resolved: Host={hostType}, UI='{pipeType}' → Global Rule='{resolvedType}' → isCircular={isCircular}");
                }
                else
                {
                    // Fallback to CONDITIONS XML if strategy not available
                isCircular = string.Equals(pipeType, "Circular", StringComparison.OrdinalIgnoreCase);
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
                    DebugLogger.Info($"[UniversalSleevePlacer] ROUND DUCT opening type from CONDITIONS XML: '{roundDuctType}' → isCircular={isCircular}");
                }
                else
                {
                    // Rectangular ducts: always rectangular opening
                    isCircular = false;
                    DebugLogger.Info($"[UniversalSleevePlacer] RECTANGULAR DUCT → isCircular=false");
                }
            }
            else
            {
                // Other categories (cable trays, accessories): rectangular
                isCircular = false;
                DebugLogger.Info($"[UniversalSleevePlacer] OTHER CATEGORY ({clashZone.MepElementCategory}) → isCircular=false");
            }
            
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
            
            DebugLogger.Info($"[UniversalSleevePlacer] Selected family: {familyName}, Type: {typeName}");
            
            return (familyName, typeName, isCircular);
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
                
                DebugLogger.Info($"[FAMILY-CACHE] ✓ Pre-cached {cachedCount} family symbols. Total cached: {_familySymbolCache.Count}");
            }
            catch (Exception ex)
            {
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
        /// Internal method to actually load family symbol from Revit (no caching)
        /// </summary>
        private FamilySymbol LoadFamilySymbolInternal(string familyName)
        {
            // Load your universal families directly
            var symbols = new FilteredElementCollector(_doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(s => s.Family.Name == familyName)
                .ToList();
            
            var symbol = symbols.FirstOrDefault();
            
            if (symbol == null)
            {
                DebugLogger.Error($"[UniversalSleevePlacer] Universal family '{familyName}' not found in project!");
                DebugLogger.Error($"[UniversalSleevePlacer] Please load the family from the Resources folder:");
                DebugLogger.Error($"[UniversalSleevePlacer] 1. Go to Insert > Load Family");
                DebugLogger.Error($"[UniversalSleevePlacer] 2. Navigate to Resources folder");
                DebugLogger.Error($"[UniversalSleevePlacer] 3. Load {familyName}.rfa");
            }
            else
            {
                DebugLogger.Info($"[UniversalSleevePlacer] Found universal family: {familyName} (ID: {symbol.Id})");
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
                DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] GetClearanceFromConditions START: category='{category}', Shape='{mepSize.Shape}', IsInsulated={mepSize.IsInsulated}\n");
                
                DebugLogger.Info($"[GetClearanceFromConditions] START: category='{category}', UI settings={_clearanceSettings.Count}, XML conditions={(_conditions != null ? "NOT NULL" : "NULL")}");
                
                // ✅ PRIORITY 1: Use UI clearance settings if available
                if (_clearanceSettings.Count > 0)
                {
                    double clearanceInMm = GetClearanceFromUISettings(category, mepSize);
                    if (clearanceInMm > 0)
                    {
                        DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] ✅ Using UI clearance: {clearanceInMm}mm for {category}\n");
                        DebugLogger.Info($"[GetClearanceFromConditions] Using UI clearance: {clearanceInMm}mm for {category}");
                        return UnitUtils.ConvertToInternalUnits(clearanceInMm, UnitTypeId.Millimeters);
                    }
                }
                
                // ✅ PRIORITY 2: Fallback to XML conditions
                if (_conditions?.ClearanceSettings != null)
                {
                    DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] Using XML clearance settings for {category}\n");
                    DebugLogger.Info($"[GetClearanceFromConditions] Using XML clearance settings for {category}");
                    return GetClearanceFromXmlConditions(category, mepSize);
                }
                
                // ✅ PRIORITY 3: Default fallback
                DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] ⚠️ No clearance settings available, using default 50mm\n");
                DebugLogger.Warning($"[GetClearanceFromConditions] No clearance settings available, using default 50mm");
                return UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters);
            }
            catch (Exception ex)
            {
                DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] ❌ ERROR in clearance calculation: {ex.Message}\n");
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
                DebugLogger.Info($"[GetClearanceFromUISettings] Checking UI settings for category: {category}");
                
                // Log all available UI clearance settings
                foreach (var kvp in _clearanceSettings)
                {
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
                        DebugLogger.Info($"[GetClearanceFromUISettings] Pipes: isInsulated={isInsulated}, key='{targetKey}', clearance={clearance}mm");
                        return clearance;
                    }
                    
                    // Fallback to generic pipe clearance
                    if (_clearanceSettings.ContainsKey("pipes_clearance"))
                    {
                        double clearance = _clearanceSettings["pipes_clearance"];
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
                        DebugLogger.Info($"[GetClearanceFromUISettings] Cable Trays: Using top clearance={clearance}mm");
                        return clearance;
                    }
                    
                    if (_clearanceSettings.ContainsKey(otherKey))
                    {
                        double clearance = _clearanceSettings[otherKey];
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
                    
                    DebugLogger.Info($"[GetClearanceFromUISettings] Ducts: Looking for key='{targetKey}', isInsulated={isInsulated}");
                    
                    if (_clearanceSettings.ContainsKey(targetKey))
                    {
                        double clearance = _clearanceSettings[targetKey];
                        DebugLogger.Info($"[GetClearanceFromUISettings] Ducts: Found key='{targetKey}', clearance={clearance}mm");
                        return clearance;
                    }
                    else
                    {
                        DebugLogger.Warning($"[GetClearanceFromUISettings] Ducts: Key '{targetKey}' not found in clearance settings!");
                        
                        // Try alternative keys
                        if (_clearanceSettings.ContainsKey("normal_clearance"))
                        {
                            double clearance = _clearanceSettings["normal_clearance"];
                            DebugLogger.Info($"[GetClearanceFromUISettings] Ducts: Using fallback 'normal_clearance' = {clearance}mm");
                            return clearance;
                        }
                        
                        if (_clearanceSettings.ContainsKey("insulated_clearance"))
                        {
                            double clearance = _clearanceSettings["insulated_clearance"];
                            DebugLogger.Info($"[GetClearanceFromUISettings] Ducts: Using fallback 'insulated_clearance' = {clearance}mm");
                            return clearance;
                        }
                    }
                }
                
                DebugLogger.Warning($"[GetClearanceFromUISettings] No UI clearance found for category: {category}");
                return 0; // Indicate no UI clearance found
            }
            catch (Exception ex)
            {
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
                DebugLogger.Info($"[GetClearanceFromXmlConditions] ClearanceSettings available: RectNormal={_conditions.ClearanceSettings.RectangularNormal}mm, RectInsulated={_conditions.ClearanceSettings.RectangularInsulated}mm");

                // Determine clearance based on category and element properties
                double clearanceInMm = 50.0; // Default fallback

                if (string.Equals(category, "Pipes", StringComparison.OrdinalIgnoreCase))
                {
                    // Pipes: Check if insulated
                    bool isInsulated = IsPipeInsulated(mepSize);
                    
                    // ⚠️ DIAGNOSTIC: Log insulation detection from XML data
                    DebugLogger.Info($"[GetClearanceFromXmlConditions] Pipes: Reading from XML - Shape='{mepSize.Shape}', IsInsulated={isInsulated}, InsulationThickness={mepSize.InsulationThickness:F6}ft");
                    DebugLogger.Info($"[GetClearanceFromXmlConditions] Pipes: isInsulated={isInsulated}, clearance={clearanceInMm}mm");
                    
                    clearanceInMm = isInsulated ? _conditions.ClearanceSettings.PipesInsulated : _conditions.ClearanceSettings.PipesNormal;
                    DebugLogger.Info($"[GetClearanceFromXmlConditions] Pipes: Final clearance={clearanceInMm}mm (insulated={isInsulated})");
                }
                else if (string.Equals(category, "Ducts", StringComparison.OrdinalIgnoreCase))
                {
                    // Ducts: Check if insulated and shape
                    bool isInsulated = IsDuctInsulated(mepSize);
                    bool isRound = string.Equals(mepSize.Shape, "Round", StringComparison.OrdinalIgnoreCase) ||
                                  string.Equals(mepSize.Shape, "Circular", StringComparison.OrdinalIgnoreCase);
                    
                    // ⚠️ DIAGNOSTIC: Log insulation detection from XML data
                    DebugLogger.Info($"[GetClearanceFromXmlConditions] Ducts: Reading from XML - Shape='{mepSize.Shape}', IsInsulated={isInsulated}, InsulationThickness={mepSize.InsulationThickness:F6}ft");
                    DebugLogger.Info($"[GetClearanceFromXmlConditions] Ducts: isInsulated={isInsulated}, isRound={isRound}, Shape='{mepSize.Shape}'");
                    
                    if (isRound)
                    {
                        clearanceInMm = isInsulated ? _conditions.ClearanceSettings.RoundInsulated : _conditions.ClearanceSettings.RoundNormal;
                        DebugLogger.Info($"[GetClearanceFromXmlConditions] Round ducts: clearance={clearanceInMm}mm");
                    }
                    else
                    {
                        clearanceInMm = isInsulated ? _conditions.ClearanceSettings.RectangularInsulated : _conditions.ClearanceSettings.RectangularNormal;
                        DebugLogger.Info($"[GetClearanceFromXmlConditions] Rectangular ducts: clearance={clearanceInMm}mm");
                    }
                }
                else if (string.Equals(category, "Cable Trays", StringComparison.OrdinalIgnoreCase))
                {
                    // Cable Trays: Use top clearance as default
                    clearanceInMm = _conditions.ClearanceSettings.CableTrayTop;
                    DebugLogger.Info($"[GetClearanceFromXmlConditions] Cable Trays: clearance={clearanceInMm}mm");
                }

                // Convert from mm to feet (Revit internal units)
                double clearanceInFeet = UnitUtils.ConvertToInternalUnits(clearanceInMm, UnitTypeId.Millimeters);
                
                DebugLogger.Info($"[GetClearanceFromXmlConditions] {category}: {clearanceInMm}mm → {clearanceInFeet:F6}ft");
                return clearanceInFeet;
            }
            catch (Exception ex)
            {
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
                    DebugLogger.Error($"[ValidateDimensions] NEGATIVE/ZERO dimensions: W={widthMm:F1}mm, H={heightMm:F1}mm, D={diameterMm:F1}mm");
                    return false;
                }
                
                // Check for oversized dimensions
                if (widthMm > MAX_SIZE_MM || heightMm > MAX_SIZE_MM || diameterMm > MAX_SIZE_MM)
                {
                    DebugLogger.Error($"[ValidateDimensions] OVERSIZED dimensions: W={widthMm:F1}mm, H={heightMm:F1}mm, D={diameterMm:F1}mm (MAX={MAX_SIZE_MM}mm)");
                    return false;
                }
                
                // Check for undersized dimensions
                if (widthMm < MIN_SIZE_MM || heightMm < MIN_SIZE_MM || diameterMm < MIN_SIZE_MM)
                {
                    DebugLogger.Error($"[ValidateDimensions] UNDERSIZED dimensions: W={widthMm:F1}mm, H={heightMm:F1}mm, D={diameterMm:F1}mm (MIN={MIN_SIZE_MM}mm)");
                    return false;
                }
                
                // Check for reasonable aspect ratio (prevent extremely thin sleeves)
                const double MAX_ASPECT_RATIO = 20.0; // Max 20:1 ratio
                double aspectRatio1 = Math.Max(widthMm, heightMm) / Math.Min(widthMm, heightMm);
                if (aspectRatio1 > MAX_ASPECT_RATIO)
                {
                    DebugLogger.Error($"[ValidateDimensions] EXTREME ASPECT RATIO: {aspectRatio1:F1}:1 (MAX={MAX_ASPECT_RATIO}:1)");
                    return false;
                }
                
                DebugLogger.Info($"[ValidateDimensions] ✓ VALID dimensions: W={widthMm:F1}mm, H={heightMm:F1}mm, D={diameterMm:F1}mm");
                return true;
            }
            catch (Exception ex)
            {
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
                // 🛡️ FAIL-SAFE: Validate dimensions before setting parameters
                if (!ValidateSleeveDimensions(finalWidth, finalHeight, finalDiameter, clashZone))
                {
                    DebugLogger.Error($"[UniversalSleevePlacer] INVALID DIMENSIONS for ClashZone {clashZone.Id}: W={finalWidth:F3}ft, H={finalHeight:F3}ft, D={finalDiameter:F3}ft");
                    ErrorCount++;
                    return; // Skip this sleeve - don't crash Revit
                }
        bool isPipe = string.Equals(clashZone.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase);
        
        // 🔍 DEBUG: Log pipe detection for floors vs walls
        var hostType = clashZone.StructuralElementType ?? "Unknown";
        DebugLogger.Info($"[SetSleeveParameters] PIPE DETECTION DEBUG: Host={hostType}, Category={clashZone.MepElementCategory}, isPipe={isPipe}");
        
        // Apply rounding to nearest 5mm if setting is enabled
                var (roundedWidth, roundedHeight) = OpeningSettingsHelper.RoundDimensionsToNearest5mm(finalWidth, finalHeight);
                var roundedDiameter = OpeningSettingsHelper.RoundDiameterToNearest5mm(finalDiameter);
        
        // Log what we're setting (with millimeter conversions for readability)
        try
        {
            double diamMm = UnitUtils.ConvertFromInternalUnits(roundedDiameter, UnitTypeId.Millimeters);
            double widthMm = UnitUtils.ConvertFromInternalUnits(roundedWidth, UnitTypeId.Millimeters);
            double heightMm = UnitUtils.ConvertFromInternalUnits(roundedHeight, UnitTypeId.Millimeters);
            DebugLogger.Info($"[PARAM-SET] Sleeve {sleeveInstance.Id.IntegerValue}: isPipe={isPipe}, Shape={mepSize.Shape}, roundedDia={diamMm:F1}mm, W={widthMm:F1}mm, H={heightMm:F1}mm\n");
        }
        catch { }
                
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
                var p = sleeveInstance.LookupParameter(name);
                if (p != null && !p.IsReadOnly)
                {
                    p.Set(roundedDiameter);
                    double diamMm = UnitUtils.ConvertFromInternalUnits(roundedDiameter, UnitTypeId.Millimeters);
                    DebugLogger.Info($"[UniversalSleevePlacer] Set '{name}' = {roundedDiameter:F6} ft ({diamMm:F1}mm) on sleeve {sleeveInstance.Id}");
                    try
                    {
                        DebugLogger.Info($"[PARAM-SUCCESS] Sleeve {sleeveInstance.Id.IntegerValue}: Set '{name}' = {diamMm:F1}mm\n");
                    }
                    catch { }
                    setOk = true;
                    break;
                }
            }
            
            if (!setOk)
                    {
                        // Fallback to Width/Height for circular
                        sleeveInstance.LookupParameter("Width")?.Set(roundedDiameter);
                        sleeveInstance.LookupParameter("Height")?.Set(roundedDiameter);
                double diamMm = UnitUtils.ConvertFromInternalUnits(roundedDiameter, UnitTypeId.Millimeters);
                DebugLogger.Info($"[UniversalSleevePlacer] Fallback set Width/Height = {diamMm:F1}mm (circular) on sleeve {sleeveInstance.Id}");
                    }
                }
                else
                {
                    // Rectangular
                    sleeveInstance.LookupParameter("Width")?.Set(roundedWidth);
                    sleeveInstance.LookupParameter("Height")?.Set(roundedHeight);
            double widthMm = UnitUtils.ConvertFromInternalUnits(roundedWidth, UnitTypeId.Millimeters);
            double heightMm = UnitUtils.ConvertFromInternalUnits(roundedHeight, UnitTypeId.Millimeters);
            DebugLogger.Info($"[UniversalSleevePlacer] Set Width={widthMm:F1}mm, Height={heightMm:F1}mm (rectangular)");
        }
        
        // ⚠️ FLOOR FIX: For duct sleeves on floors, rotate orientation by 90 degrees and swap width/height
        // NOTE: Cable trays should NOT have width/height swapped - they maintain their original orientation
        bool isFloorHost = string.Equals(clashZone.StructuralElementType, "Floor", StringComparison.OrdinalIgnoreCase) ||
                          string.Equals(clashZone.StructuralElementType, "Floors", StringComparison.OrdinalIgnoreCase);
        bool isDuct = string.Equals(clashZone.MepElementCategory, "Ducts", StringComparison.OrdinalIgnoreCase) ||
                      string.Equals(clashZone.MepElementCategory, "Duct Accessories", StringComparison.OrdinalIgnoreCase);
        bool isCableTray = string.Equals(clashZone.MepElementCategory, "Cable Trays", StringComparison.OrdinalIgnoreCase) ||
                          string.Equals(clashZone.MepElementCategory, "Cable Tray Fittings", StringComparison.OrdinalIgnoreCase);

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
                
                DebugLogger.Info($"[UniversalSleevePlacer] FLOOR DUCT SWAP: Height({UnitUtils.ConvertFromInternalUnits(roundedHeight, UnitTypeId.Millimeters):F1}mm) > Width({UnitUtils.ConvertFromInternalUnits(tempWidth, UnitTypeId.Millimeters):F1}mm) - Swapped to make longer dimension the width");
            }
            else
            {
                // Width is already longer - no swap needed
                DebugLogger.Info($"[UniversalSleevePlacer] FLOOR DUCT NO SWAP: Width({UnitUtils.ConvertFromInternalUnits(roundedWidth, UnitTypeId.Millimeters):F1}mm) >= Height({UnitUtils.ConvertFromInternalUnits(roundedHeight, UnitTypeId.Millimeters):F1}mm) - Longer dimension already width");
            }
            
            // Update the sleeve parameters with correct values
            sleeveInstance.LookupParameter("Width")?.Set(roundedWidth);
            sleeveInstance.LookupParameter("Height")?.Set(roundedHeight);
            
            DebugLogger.Info($"[UniversalSleevePlacer] FLOOR DUCT FINAL: Width={UnitUtils.ConvertFromInternalUnits(roundedWidth, UnitTypeId.Millimeters):F1}mm, Height={UnitUtils.ConvertFromInternalUnits(roundedHeight, UnitTypeId.Millimeters):F1}mm (NO ROTATION)");
        }
        else if (isFloorHost && isCableTray && !treatAsCircular)
        {
            // Cable trays on floors: NO width/height swap - maintain original orientation
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
                DebugLogger.Info($"[UniversalSleevePlacer] X-FRAMING SWAP: Width={UnitUtils.ConvertFromInternalUnits(roundedWidth, UnitTypeId.Millimeters):F1}mm, Height={UnitUtils.ConvertFromInternalUnits(roundedHeight, UnitTypeId.Millimeters):F1}mm");
            }
            else if (hostOrientation == "Y")
            {
                // Y-framing: swap width and depth
                double tempWidth = roundedWidth;
                roundedWidth = roundedHeight;  // Use height as width
                roundedHeight = tempWidth;     // Use original width as height
                DebugLogger.Info($"[UniversalSleevePlacer] Y-FRAMING SWAP: Width={UnitUtils.ConvertFromInternalUnits(roundedWidth, UnitTypeId.Millimeters):F1}mm, Height={UnitUtils.ConvertFromInternalUnits(roundedHeight, UnitTypeId.Millimeters):F1}mm");
            }
        }
        
                var depthParam = sleeveInstance.LookupParameter("Depth");
                var wallWidthParam = sleeveInstance.LookupParameter("Wall Width");
        
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
        
        try
        {
            DebugLogger.Info($"[DEPTH-CHECK] Sleeve {sleeveInstance.Id.IntegerValue}: Host={clashZone.StructuralElementType}, isWall={isWallHost}, isFraming={isFramingHost}, HasDepth={depthParam != null}, DepthRO={depthParam?.IsReadOnly}, HasWallWidth={wallWidthParam != null}, WallWidthRO={wallWidthParam?.IsReadOnly}, StructThickness={thicknessMm:F1}mm\n");
        }
        catch { }
        
        bool depthSetSuccess = false;
        
        if (isWallHost && wallWidthParam != null && !wallWidthParam.IsReadOnly)
        {
            // Wall host: use Wall Width parameter
                    wallWidthParam.Set(thickness);
            DebugLogger.Info($"[UniversalSleevePlacer] WALL: Set Wall Width = {thicknessMm:F1}mm");
            try
            {
                DebugLogger.Info($"[DEPTH-SET] Sleeve {sleeveInstance.Id.IntegerValue}: WALL - Set Wall Width = {thicknessMm:F1}mm ✓\n");
            }
            catch { }
            depthSetSuccess = true;
                }
                else if (depthParam != null && !depthParam.IsReadOnly)
                {
            // Floor/Framing host: use Depth parameter
                    depthParam.Set(thickness);
            DebugLogger.Info($"[UniversalSleevePlacer] {(isFramingHost ? "FRAMING" : "FLOOR")}: Set Depth = {thicknessMm:F1}mm");
            try
            {
                // Verify the parameter was set correctly
                double depthRead = depthParam.AsDouble();
                double depthReadMm = UnitUtils.ConvertFromInternalUnits(depthRead, UnitTypeId.Millimeters);
                bool verified = Math.Abs(depthRead - thickness) < 0.0001;
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log",
                    $"[DEPTH-SET] Sleeve {sleeveInstance.Id.IntegerValue}: {(isFramingHost ? "FRAMING" : (isWallHost ? "WALL" : "FLOOR"))} - Set Depth = {depthReadMm:F1}mm, Verified={verified} {(verified ? "✓" : "✗")}\n");
            }
            catch { }
            depthSetSuccess = true;
        }
        else
        {
            // Try type parameter as fallback
            var typeDepthParam = sleeveInstance.Symbol?.LookupParameter("Depth");
            if (typeDepthParam != null && !typeDepthParam.IsReadOnly)
            {
                typeDepthParam.Set(thickness);
                DebugLogger.Info($"[UniversalSleevePlacer] Set TYPE Depth = {thicknessMm:F1}mm (instance param not writable)");
                try
                {
                    _doc.Regenerate();
                    DebugLogger.Info($"[DEPTH-SET] Sleeve {sleeveInstance.Id.IntegerValue}: TYPE Depth = {thicknessMm:F1}mm (regenerated) ✓\n");
                }
                catch { }
                depthSetSuccess = true;
            }
        }
        
        if (!depthSetSuccess)
        {
            DebugLogger.Warning($"[UniversalSleevePlacer] WARNING: Could not find writable Depth or Wall Width parameter on sleeve {sleeveInstance.Id}");
            try
            {
                DebugLogger.Info($"[DEPTH-FAIL] Sleeve {sleeveInstance.Id.IntegerValue}: No writable depth parameter found! ✗\n");
            }
            catch { }
        }
        
        // ✅ BOTTOM OF OPENING: Calculate and set for rectangular openings on walls and framing
        // Formula: Bottom of Opening = Elevation from Level - (Height / 2)
        // This calculates the bottom edge of the opening for scheduling purposes
        if (!treatAsCircular && (isWallHost || isFramingHost))
        {
            try
            {
                // Try multiple parameter name variations for "Elevation from Level"
                var elevationFromLevelParam = sleeveInstance.LookupParameter("Elevation from Level") 
                                           ?? sleeveInstance.LookupParameter("Schedule Level Elevation")
                                           ?? sleeveInstance.LookupParameter("Elevation from Level Offset");
                
                var heightParam = sleeveInstance.LookupParameter("Height");
                var bottomOfOpeningParam = sleeveInstance.LookupParameter("Bottom of Opening");
                
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
                        
                        double elevationMm = UnitUtils.ConvertFromInternalUnits(elevationFromLevel, UnitTypeId.Millimeters);
                        double heightMm = UnitUtils.ConvertFromInternalUnits(height, UnitTypeId.Millimeters);
                        double bottomMm = UnitUtils.ConvertFromInternalUnits(bottomOfOpening, UnitTypeId.Millimeters);
                        
                        DebugLogger.Info($"[UniversalSleevePlacer] ✅ Set Bottom of Opening = {bottomMm:F1}mm (Elevation={elevationMm:F1}mm - Height/2={heightMm/2:F1}mm) for {clashZone.StructuralElementType}");
                    }
                    else
                    {
                        DebugLogger.Warning($"[UniversalSleevePlacer] Invalid values for Bottom of Opening calculation: Elevation={elevationFromLevel:F3}ft, Height={height:F3}ft");
                    }
                }
                else
                {
                    var elevationFound = elevationFromLevelParam != null;
                    var heightFound = heightParam != null;
                    var bottomFound = bottomOfOpeningParam != null && !bottomOfOpeningParam.IsReadOnly;
                    DebugLogger.Warning($"[UniversalSleevePlacer] Cannot set Bottom of Opening: Elevation from Level={elevationFound}, Height={heightFound}, Bottom of Opening={bottomFound}");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Warning($"[UniversalSleevePlacer] Error calculating Bottom of Opening: {ex.Message}");
            }
        }
        
        // Set MEP metadata parameters
                var mepElementIdParam = sleeveInstance.LookupParameter("MEP_ElementId");
                if (mepElementIdParam != null && !mepElementIdParam.IsReadOnly)
                {
                    mepElementIdParam.Set(clashZone.MepElementId.IntegerValue);
            DebugLogger.Info($"[UniversalSleevePlacer] Set MEP_ElementId = {clashZone.MepElementId.IntegerValue}");
                }

                sleeveInstance.LookupParameter("MEP_UniqueId")?.Set(clashZone.MepElementUniqueId);
                sleeveInstance.LookupParameter("MEP_Size")?.Set(clashZone.MepElementFormattedSize);
                sleeveInstance.LookupParameter("System_Abbreviation")?.Set(clashZone.MepElementSystemAbbreviation);
                sleeveInstance.LookupParameter("MEP_Count")?.Set(1);  // Individual sleeve
                
                // ✅ CRITICAL FIX: Set MEP_Category parameter for clustering
                var mepCategoryParam = sleeveInstance.LookupParameter("MEP_Category");
                if (mepCategoryParam != null && !mepCategoryParam.IsReadOnly)
                {
                    mepCategoryParam.Set(clashZone.MepElementCategory);
                    DebugLogger.Info($"[UniversalSleevePlacer] Set MEP_Category = '{clashZone.MepElementCategory}' for clustering");
                    
                    // ✅ CRITICAL LOGGING: Log MEP_Category parameter setting
                    DebugLogger.Info($"[MEP-CATEGORY-SET] {DateTime.Now:HH:mm:ss.fff} - Sleeve {sleeveInstance.Id.IntegerValue}: Set MEP_Category = '{clashZone.MepElementCategory}'\n");
                }
                else
                {
                    DebugLogger.Warning($"[UniversalSleevePlacer] MEP_Category parameter not found or read-only on sleeve {sleeveInstance.Id}");
                    
                    // ✅ CRITICAL LOGGING: Log MEP_Category parameter failure
                    DebugLogger.Info($"[MEP-CATEGORY-FAILED] {DateTime.Now:HH:mm:ss.fff} - Sleeve {sleeveInstance.Id.IntegerValue}: MEP_Category parameter not found or read-only\n");
                }

                // ⚠️ CRITICAL FIX: Transfer HostParameterValues from XML intersection data to sleeve parameters
                if (clashZone.HostParameterValues != null && clashZone.HostParameterValues.Count > 0)
                {
                    DebugLogger.Info($"[UniversalSleevePlacer] Transferring {clashZone.HostParameterValues.Count} host parameters from XML intersection data");
                    
                    foreach (var hostParam in clashZone.HostParameterValues)
                    {
                        try
                        {
                            var param = sleeveInstance.LookupParameter(hostParam.Key);
                            if (param != null && !param.IsReadOnly)
                            {
                                // Handle different parameter storage types
                                if (param.StorageType == StorageType.String)
                                {
                                    param.Set(hostParam.Value);
                                    DebugLogger.Info($"[UniversalSleevePlacer] Set host parameter '{hostParam.Key}' = '{hostParam.Value}' (string)");
                                }
                                else if (param.StorageType == StorageType.Integer)
                                {
                                    if (int.TryParse(hostParam.Value, out int intValue))
                                    {
                                        param.Set(intValue);
                                        DebugLogger.Info($"[UniversalSleevePlacer] Set host parameter '{hostParam.Key}' = {intValue} (integer)");
                                    }
                                }
                                else if (param.StorageType == StorageType.Double)
                                {
                                    if (double.TryParse(hostParam.Value, out double doubleValue))
                                    {
                                        param.Set(doubleValue);
                                        DebugLogger.Info($"[UniversalSleevePlacer] Set host parameter '{hostParam.Key}' = {doubleValue} (double)");
                                    }
                                }
                                else if (param.StorageType == StorageType.ElementId)
                                {
                                    // Handle ElementId parameters - this might need special handling
                                    DebugLogger.Warning($"[UniversalSleevePlacer] Host parameter '{hostParam.Key}' is ElementId type - skipping (value: '{hostParam.Value}')");
                                }
                                
                                // Log successful parameter transfer
                                try
                                {
                                    DebugLogger.Info($"[HOST-PARAM-SET] Sleeve {sleeveInstance.Id.IntegerValue}: Set '{hostParam.Key}' = '{hostParam.Value}' ({param.StorageType}) ✓\n");
                                }
                                catch { }
                            }
                            else
                            {
                                DebugLogger.Warning($"[UniversalSleevePlacer] Host parameter '{hostParam.Key}' not found or read-only on sleeve {sleeveInstance.Id}");
                                try
                                {
                                    DebugLogger.Info($"[HOST-PARAM-MISSING] Sleeve {sleeveInstance.Id.IntegerValue}: Parameter '{hostParam.Key}' not found or read-only ✗\n");
                                }
                                catch { }
                            }
                        }
                        catch (Exception paramEx)
                        {
                            DebugLogger.Warning($"[UniversalSleevePlacer] Error setting host parameter '{hostParam.Key}' = '{hostParam.Value}': {paramEx.Message}");
                            try
                            {
                                DebugLogger.Info($"[HOST-PARAM-ERROR] Sleeve {sleeveInstance.Id.IntegerValue}: '{hostParam.Key}' = '{hostParam.Value}' - {paramEx.Message} ✗\n");
                            }
                            catch { }
                        }
                    }
                }
                else
                {
                    DebugLogger.Warning($"[UniversalSleevePlacer] No HostParameterValues found in ClashZone {clashZone.Id} - host parameters not transferred");
                    try
                    {
                        DebugLogger.Info($"[HOST-PARAM-EMPTY] Sleeve {sleeveInstance.Id.IntegerValue}: No HostParameterValues in ClashZone {clashZone.Id} ✗\n");
                    }
                    catch { }
                }

        DebugLogger.Info($"[UniversalSleevePlacer] ✓ Set all parameters for sleeve {sleeveInstance.Id}");
            }
            catch (Exception ex)
            {
        DebugLogger.Warning($"[UniversalSleevePlacer] Error setting parameters: {ex.Message}\n{ex.StackTrace}");
        try
        {
            DebugLogger.Info($"[PARAM-ERROR] Sleeve {sleeveInstance.Id.IntegerValue}: {ex.Message}\n");
        }
        catch { }
    }
}

        /// <summary>
        /// Set sleeve metadata parameters for fast parameter transfer
        /// </summary>
        private void SetSleeveMetadata(FamilyInstance sleeveInstance, string category)
        {
            try
            {
                // Set Filter Name based on category
                string filterName = GetFilterNameForCategory(category);
                var filterNameParam = sleeveInstance.LookupParameter("Filter Name");
                if (filterNameParam != null && !filterNameParam.IsReadOnly)
                {
                    filterNameParam.Set(filterName);
                    DebugLogger.Info($"[SetSleeveMetadata] Set Filter Name = '{filterName}' for sleeve {sleeveInstance.Id}");
                }
                else
                {
                    DebugLogger.Warning($"[SetSleeveMetadata] Filter Name parameter not found or read-only on sleeve {sleeveInstance.Id}");
                }
                
                // Set Instance ID
                var instanceIdParam = sleeveInstance.LookupParameter("Sleeve Instance ID");
                if (instanceIdParam != null && !instanceIdParam.IsReadOnly)
                {
                    instanceIdParam.Set(sleeveInstance.Id.IntegerValue);
                    DebugLogger.Info($"[SetSleeveMetadata] Set Sleeve Instance ID = {sleeveInstance.Id.IntegerValue} for sleeve {sleeveInstance.Id}");
                }
                else
                {
                    DebugLogger.Warning($"[SetSleeveMetadata] Sleeve Instance ID parameter not found or read-only on sleeve {sleeveInstance.Id}");
                }
                
                // Log to debug file
                try
                {
                    DebugLogger.Info($"[SLEEVE_METADATA] Sleeve {sleeveInstance.Id.IntegerValue}: Filter='{filterName}', InstanceID={sleeveInstance.Id.IntegerValue} ✓\n");
                }
                catch { }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[SetSleeveMetadata] Error setting metadata for sleeve {sleeveInstance.Id}: {ex.Message}");
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
                            
                            DebugLogger.Info($"[COLLECT] Sleeve {sleeve.Id.IntegerValue}: Host={sleeveData.HostType}, Orientation={sleeveData.Orientation}, W={sleeveData.Width:F3}, H={sleeveData.Height:F3}, D={sleeveData.Depth:F3}\n");
                        }
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Info($"[ERROR] Failed to collect data for sleeve {sleeve.Id.IntegerValue}: {ex.Message}\n");
                    }
                }
            }
            catch (Exception ex)
            {
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
                    DebugLogger.Info($"[XML-IMMEDIATE-UPDATE] ❌ Filters directory does not exist: {filtersDirectory}\n");
                    return;
                }

                // 🎯 CRITICAL FIX: Use category-based file matching to target the correct XML file
                var targetFileName = GetFilterNameForCategory(clashZone.MepElementCategory);
                
                // ✅ CRITICAL LOGGING: Log XML file details
                DebugLogger.Info($"[XML-FILE-SEARCH] {DateTime.Now:HH:mm:ss.fff} - Looking for zone {clashZone.Id} in category file: {targetFileName}\n");
                DebugLogger.Info($"[XML-FILE-SEARCH] Category: {clashZone.MepElementCategory}, TargetFileName: {targetFileName}\n");
                
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
                        DebugLogger.Info($"[XML-FILE-FOUND] {DateTime.Now:HH:mm:ss.fff} - Found target file: {fileName}\n");
                        DebugLogger.Info($"[XML-FILE-FOUND] Full path: {xmlFile}\n");
                        
                        DebugLogger.Info($"[XML-IMMEDIATE-UPDATE] Found target file: {fileName}\n");
                        break;
                    }
                }

                // If no specific file found, fall back to searching all files
                if (targetFile == null)
                {
                    DebugLogger.Info($"[XML-IMMEDIATE-UPDATE] No specific file found, searching all {xmlFiles.Length} XML files\n");
                    targetFile = xmlFiles.FirstOrDefault(f => !Path.GetFileName(f).Contains("CONDITIONS", StringComparison.OrdinalIgnoreCase));
                }

                if (targetFile == null)
                {
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
                
                if (filter?.ClashZoneStorage?.ClashZones == null)
                {
                    DebugLogger.Info($"[XML-IMMEDIATE-UPDATE] ❌ No clash zones in file: {Path.GetFileName(targetFile)}\n");
                    return;
                }

                // 🔥 TRIPLE-MATCH VERIFICATION: Match by ID AND MepElementId AND StructuralElementId
                var zone = filter.ClashZoneStorage.ClashZones.FirstOrDefault(z => 
                    z.Id == clashZone.Id && 
                    z.MepElementIdValue == clashZone.MepElementIdValue && 
                    z.StructuralElementIdValue == clashZone.StructuralElementIdValue);
                
                if (zone != null)
                {
                    DebugLogger.Info($"[XML-IMMEDIATE-UPDATE] ✅ FOUND zone {clashZone.Id} in {Path.GetFileName(targetFile)}\n");
                    
                    // Log BEFORE values
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
                    // Coordinates will be updated later using correct timing
                    
                    // Log AFTER values
                    DebugLogger.Info($"[XML-IMMEDIATE-UPDATE] AFTER: W={UnitUtils.ConvertFromInternalUnits(zone.SleeveWidth, UnitTypeId.Millimeters):F1}mm, H={UnitUtils.ConvertFromInternalUnits(zone.SleeveHeight, UnitTypeId.Millimeters):F1}mm, ActiveDocX={zone.SleevePlacementPointActiveDocumentX:F3}\n");
                    
                    // Save the file immediately
                    filter.LastModified = DateTime.Now;
                    using (var writer = new StreamWriter(targetFile))
                    {
                        serializer.Serialize(writer, filter);
                    }
                    
                    // ✅ CRITICAL LOGGING: Log successful XML save with all details
                    DebugLogger.Info($"[XML-SAVE-SUCCESS] {DateTime.Now:HH:mm:ss.fff} - Successfully saved to {Path.GetFileName(targetFile)}\n");
                    DebugLogger.Info($"[XML-SAVE-SUCCESS] Zone: {clashZone.Id}, SleeveInstanceId: {zone.SleeveInstanceId}\n");
                    DebugLogger.Info($"[XML-SAVE-SUCCESS] Coordinates: X={zone.SleevePlacementPointActiveDocumentX:F3}, Y={zone.SleevePlacementPointActiveDocumentY:F3}, Z={zone.SleevePlacementPointActiveDocumentZ:F3}\n");
                    DebugLogger.Info($"[XML-SAVE-SUCCESS] Dimensions: W={UnitUtils.ConvertFromInternalUnits(zone.SleeveWidth, UnitTypeId.Millimeters):F1}mm, H={UnitUtils.ConvertFromInternalUnits(zone.SleeveHeight, UnitTypeId.Millimeters):F1}mm\n");
                    
                    DebugLogger.Info($"[XML-IMMEDIATE-UPDATE] ✅ SUCCESS: Updated {Path.GetFileName(targetFile)} with values for zone {clashZone.Id}\n");
                }
                else
                {
                    DebugLogger.Info($"[XML-IMMEDIATE-UPDATE] ❌ Zone {clashZone.Id} NOT found in {Path.GetFileName(targetFile)} (checked {filter.ClashZoneStorage.ClashZones.Count} zones)\n");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Info($"[XML-IMMEDIATE-UPDATE-ERROR] Error updating XML: {ex.Message}\n");
            }
        }

        /// <summary>
        /// CRITICAL FIX: Save XML files with updated SleeveInstanceId values
        /// This ensures parameter transfer can find the sleeves in the XML
        /// </summary>
        private void SaveUpdatedXmlFiles(List<ClashZone> updatedClashZones = null)
        {
            try
            {
                DebugLogger.Info("[UniversalSleevePlacer] Saving updated XML files with SleeveInstanceId values...");
                DebugLogger.Info($"[{DateTime.Now}] [XML_SAVE] Starting SaveUpdatedXmlFiles method with {updatedClashZones?.Count ?? 0} updated clash zones\n");
                
                var filtersDirectory = ProjectPathService.GetFiltersDirectory(_doc);
                
                if (!Directory.Exists(filtersDirectory))
                {
                    DebugLogger.Warning($"[UniversalSleevePlacer] Filters directory not found: {filtersDirectory}");
                    return;
                }
                
                // Get all XML files
                var xmlFiles = Directory.GetFiles(filtersDirectory, "*.xml");
                DebugLogger.Info($"[UniversalSleevePlacer] Found {xmlFiles.Length} XML files to check");
                
                foreach (var xmlFile in xmlFiles)
                {
                    try
                    {
                        // Skip CONDITIONS files
                        if (Path.GetFileName(xmlFile).Contains("CONDITIONS", StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }
                        
                        // ✅ CRITICAL FIX: Load XML but ensure we don't overwrite current values
                        var serializer = new System.Xml.Serialization.XmlSerializer(typeof(Models.OpeningFilter));
                        Models.OpeningFilter filter;
                        
                        using (var reader = new StreamReader(xmlFile))
                        {
                            filter = (Models.OpeningFilter)serializer.Deserialize(reader);
                        }
                        
                        if (filter?.ClashZoneStorage?.ClashZones == null)
                        {
                            continue;
                        }
                        
                        // ✅ CRITICAL FIX: Merge in-memory updated clash zones with XML data
                        if (updatedClashZones != null && updatedClashZones.Count > 0)
                        {
                            DebugLogger.Info($"[UniversalSleevePlacer] Merging {updatedClashZones.Count} in-memory clash zones into XML file");
                            
                            foreach (var updatedZone in updatedClashZones)
                            {
                                // Find matching zone in the loaded XML by ID
                                var matchingZone = filter.ClashZoneStorage.ClashZones.FirstOrDefault(z => z.Id == updatedZone.Id);
                                if (matchingZone != null)
                                {
                                    // ✅ CRITICAL: Copy updated values from in-memory zone to XML zone
                                    // This ensures IsResolved, SleeveInstanceId, and other flags are preserved
                                    matchingZone.IsResolved = updatedZone.IsResolved;
                                    matchingZone.SleeveInstanceId = updatedZone.SleeveInstanceId;
                                    matchingZone.SleeveFamilyName = updatedZone.SleeveFamilyName;
                                    matchingZone.IsClusterResolved = updatedZone.IsClusterResolved;
                                    matchingZone.ClusterSleeveInstanceId = updatedZone.ClusterSleeveInstanceId;
                                    
                                    // Copy sleeve data
                                    matchingZone.SleeveWidth = updatedZone.SleeveWidth;
                                    matchingZone.SleeveHeight = updatedZone.SleeveHeight;
                                    matchingZone.SleeveDiameter = updatedZone.SleeveDiameter;
                                    matchingZone.SleevePlacementPointX = updatedZone.SleevePlacementPointX;
                                    matchingZone.SleevePlacementPointY = updatedZone.SleevePlacementPointY;
                                    matchingZone.SleevePlacementPointZ = updatedZone.SleevePlacementPointZ;
                                    matchingZone.SleevePlacementPointActiveDocumentX = updatedZone.SleevePlacementPointActiveDocumentX;
                                    matchingZone.SleevePlacementPointActiveDocumentY = updatedZone.SleevePlacementPointActiveDocumentY;
                                    matchingZone.SleevePlacementPointActiveDocumentZ = updatedZone.SleevePlacementPointActiveDocumentZ;
                                    
                                    DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] [MERGE] Zone {updatedZone.Id}: IsResolved={updatedZone.IsResolved}, SleeveInstanceId={updatedZone.SleeveInstanceId}\n");
                                }
                            }
                        }
                        
                        bool hasUpdates = false;
                        foreach (var zone in filter.ClashZoneStorage.ClashZones)
                        {
                            // Check if this zone has a valid SleeveInstanceId (not -1) OR has been resolved
                            if (zone.SleeveInstanceId > 0 || zone.IsResolved == true || zone.IsClusterResolved == true)
                            {
                                hasUpdates = true;
                                DebugLogger.Info($"[UniversalSleevePlacer] Found valid SleeveInstanceId {zone.SleeveInstanceId} in {Path.GetFileName(xmlFile)}");
                                
                                // 🔥 DEBUG: Log dimension values during XML save
                                DebugLogger.Info($"[XML-SAVE-CHECK] Zone {zone.Id}: SleeveInstanceId={zone.SleeveInstanceId}, W={UnitUtils.ConvertFromInternalUnits(zone.SleeveWidth, UnitTypeId.Millimeters):F1}mm, H={UnitUtils.ConvertFromInternalUnits(zone.SleeveHeight, UnitTypeId.Millimeters):F1}mm, D={UnitUtils.ConvertFromInternalUnits(zone.SleeveDiameter, UnitTypeId.Millimeters):F1}mm\n");
                                
                                // ✅ CRITICAL FIX: The issue is that XML values are 0, but we need to ensure they're saved correctly
                                // The problem is that the XML serialization is working, but the values are being reset somewhere
                                // Let's add debug logging to track this issue
                                if (zone.SleeveWidth == 0 || zone.SleeveHeight == 0)
                                {
                                    DebugLogger.Info($"[XML-SAVE-ISSUE] Zone {zone.Id}: Dimensions are 0 - W={UnitUtils.ConvertFromInternalUnits(zone.SleeveWidth, UnitTypeId.Millimeters):F1}mm, H={UnitUtils.ConvertFromInternalUnits(zone.SleeveHeight, UnitTypeId.Millimeters):F1}mm\n");
                                }
                            }
                        }
                        
                        // Save the file if it has updates
                        if (hasUpdates)
                        {
                            // ⚠️ SAFETY MEASURE: Create backup before saving (if validation enabled)
                            if (OptimizationFlags.UseXmlValidation)
                            {
                                CreateXmlBackup(xmlFile);
                            }
                            
                            // ⚠️ CRITICAL: Log flag states BEFORE XML file save
                            DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] [XML-FILE-SAVE-BEFORE] Saving {Path.GetFileName(xmlFile)} with {filter.ClashZoneStorage.ClashZones.Count} clash zones\n");
                            
                            // Log first 3 clash zones for verification
                            foreach (var zone in filter.ClashZoneStorage.ClashZones.Take(3))
                            {
                                DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] [XML-FILE-SAVE-BEFORE] ClashZone {zone.Id}: IsResolved={zone.IsResolved}, IsClusterResolved={zone.IsClusterResolved}\n");
                                
                                // 🔥 DEBUG: Log ActiveDocument coordinates before XML save
                                DebugLogger.Info($"[XML-SAVE-BEFORE] Zone {zone.Id}: ActiveDoc={zone.SleevePlacementPointActiveDocument}, X={zone.SleevePlacementPointActiveDocumentX:F3}, Y={zone.SleevePlacementPointActiveDocumentY:F3}, Z={zone.SleevePlacementPointActiveDocumentZ:F3}\n");
                                DebugLogger.Info($"[XML-SAVE-BEFORE] Zone {zone.Id}: W={UnitUtils.ConvertFromInternalUnits(zone.SleeveWidth, UnitTypeId.Millimeters):F1}mm, H={UnitUtils.ConvertFromInternalUnits(zone.SleeveHeight, UnitTypeId.Millimeters):F1}mm\n");
                            }
                            
                            // Store expected counts for validation
                            int expectedUpdatedCount = updatedClashZones?.Count ?? 0;
                            var expectedUpdatedIds = updatedClashZones?.Select(cz => cz.Id).ToList() ?? new List<Guid>();
                            
                            filter.LastModified = DateTime.Now;
                            
                            using (var writer = new StreamWriter(xmlFile))
                            {
                                serializer.Serialize(writer, filter);
                            }
                            
                            // ⚠️ SAFETY MEASURE: Validate XML save completed correctly (if validation enabled)
                            if (OptimizationFlags.UseXmlValidation && expectedUpdatedCount > 0)
                            {
                                if (!ValidateXmlSave(xmlFile, expectedUpdatedIds, expectedUpdatedCount))
                                {
                                    // Validation failed - restore from backup
                                    DebugLogger.Error($"[XML-VERIFY] Validation failed for {Path.GetFileName(xmlFile)} - restoring from backup");
                                    RestoreXmlFromBackup(xmlFile);
                                    throw new InvalidOperationException($"XML save verification failed for {Path.GetFileName(xmlFile)} - restored from backup");
                                }
                                else
                                {
                                    DebugLogger.Info($"[XML-VERIFY] ✓ Validation passed for {Path.GetFileName(xmlFile)}: {expectedUpdatedCount} zones updated correctly");
                                }
                            }
                            
                            // ⚠️ CRITICAL: Log flag states AFTER XML file save
                            DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] [XML-FILE-SAVE-AFTER] Saved {Path.GetFileName(xmlFile)} successfully\n");
                            
                            DebugLogger.Info($"[UniversalSleevePlacer] ✓ Saved updated XML file: {Path.GetFileName(xmlFile)}");
                        }
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Error($"[UniversalSleevePlacer] Error processing XML file {Path.GetFileName(xmlFile)}: {ex.Message}");
                    }
                }
                
                DebugLogger.Info("[UniversalSleevePlacer] XML file saving complete");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UniversalSleevePlacer] Error saving XML files: {ex.Message}");
                SafeFileLogger.SafeAppendText("xml_save_errors.log", 
                    $"Error saving XML files: {ex.Message}\nStack trace: {ex.StackTrace}");
                throw; // Re-throw to prevent silent failures
            }
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
                    DebugLogger.Info($"[XML-BACKUP] Created backup: {Path.GetFileName(backupPath)}");
                }
            }
            catch (Exception ex)
            {
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
                    DebugLogger.Info($"[XML-RESTORE] ✓ Restored {Path.GetFileName(xmlFilePath)} from backup");
                    SafeFileLogger.SafeAppendText("xml_restore.log", 
                        $"Restored {Path.GetFileName(xmlFilePath)} from backup at {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                }
                else
                {
                    DebugLogger.Warning($"[XML-RESTORE] No backup found for {Path.GetFileName(xmlFilePath)}");
                }
            }
            catch (Exception ex)
            {
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
                
                if (verification?.ClashZoneStorage?.ClashZones == null)
                {
                    DebugLogger.Warning($"[XML-VERIFY] No clash zones found in saved file {Path.GetFileName(xmlFilePath)}");
                    return false;
                }
                
                // Count how many expected zones were actually updated
                int actualUpdated = 0;
                foreach (var expectedId in expectedUpdatedIds)
                {
                    var savedZone = verification.ClashZoneStorage.ClashZones.FirstOrDefault(z => z.Id == expectedId);
                    if (savedZone != null && (savedZone.SleeveInstanceId > 0 || savedZone.IsResolved))
                    {
                        actualUpdated++;
                    }
                }
                
                // Allow 5% tolerance for minor discrepancies
                double successRate = expectedCount > 0 ? (double)actualUpdated / expectedCount : 0;
                bool isValid = successRate >= 0.95;
                
                if (!isValid)
                {
                    DebugLogger.Error($"[XML-VERIFY] Validation failed: Expected {expectedCount} updates, got {actualUpdated} ({successRate:P1}). Threshold: 95%");
                }
                
                return isValid;
            }
            catch (Exception ex)
            {
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
                                    DebugLogger.Info($"[UniversalSleevePlacer] VERTICAL CABLETRAY ON FLOOR: MEP orientation is X - rotating sleeve 90°");
                                }
                                else
                                {
                                    rotationAngle = 0.0; // No rotation needed
                                    DebugLogger.Info($"[UniversalSleevePlacer] VERTICAL CABLETRAY ON FLOOR: MEP orientation is Y - no rotation needed");
                                }
                            }
                            else
                            {
                                // ⚠️ STANDARD LOGIC FOR DUCTS/PIPES
                                if (mepOrientation == "Y")
                                {
                                    rotationAngle = Math.PI / 2; // 90 degrees
                                    DebugLogger.Info($"[UniversalSleevePlacer] VERTICAL {clashZone.MepElementCategory.ToUpper()} ON FLOOR: MEP orientation is Y - rotating sleeve 90°");
                                }
                                else
                                {
                                    rotationAngle = 0.0; // No rotation needed
                                    DebugLogger.Info($"[UniversalSleevePlacer] VERTICAL {clashZone.MepElementCategory.ToUpper()} ON FLOOR: MEP orientation is X - no rotation needed");
                                }
                            }
                            
                            double rotationAngleDegrees = rotationAngle * 180 / Math.PI;
                            
                            Line rotationAxis = Line.CreateBound(loc.Point, loc.Point + XYZ.BasisZ);
                            ElementTransformUtils.RotateElement(_doc, sleeveInstance.Id, rotationAxis, rotationAngle);
                            
                    DebugLogger.Info($"[UniversalSleevePlacer] FLOOR: Rotated sleeve {rotationAngleDegrees:F1}° based on MEP orientation ({mepOrientation})");
                    try
                    {
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
                DebugLogger.Info($"[ORIENT-INPUT] Sleeve {sleeveInstance.Id.IntegerValue}: Cat={clashZone.MepElementCategory}, Host={clashZone.StructuralElementType}, MEP=({mepOrientation?.X:F3},{mepOrientation?.Y:F3}), StructN=({structuralNormal?.X:F3},{structuralNormal?.Y:F3})\n");
                
                // DEBUG: Log all ClashZone properties to trace XML loading issue
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
                            DebugLogger.Info($"[WALL-LOGIC-REACHED] Sleeve {sleeveInstance.Id.IntegerValue}: Reached wall direction logic block ✓\n");
                        }
                        catch { }
                        
                        // ⚠️ CRITICAL OPTIMIZATION: Use pre-calculated wall direction from XML ⚠️
                        // NO EXPENSIVE REVIT API CALLS - wall direction calculated during refresh
                        bool allowRotate = false;
                        
                        // DIAGNOSTIC: Log ClashZone wall direction data availability
                        try
                        {
                            System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log",
                                $"[DIAGNOSTIC] Sleeve {sleeveInstance.Id.IntegerValue}: WallDirection={clashZone.WallDirection != null}, WallDirectionType='{clashZone.WallDirectionType ?? "NULL"}', WallDirectionZero={clashZone.WallDirection == XYZ.Zero}\n");
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
                            DebugLogger.Info($"[UniversalSleevePlacer] HostOrientation=X: Will apply +90° rotation");
                        }
                        else if (hostOrientation == "Y")
                        {
                            needsRotation = false; // Y-orientation works naturally with LEFT view families
                            DebugLogger.Info($"[UniversalSleevePlacer] HostOrientation=Y: No rotation needed");
                        }
                        else
                        {
                            // Fallback: Use wall direction type if HostOrientation not available
                            if (wallDirectionType == "X-WALL")
                            {
                                needsRotation = true;
                                DebugLogger.Warning($"[UniversalSleevePlacer] Fallback: X-WALL detected - will apply +90° rotation");
                            }
                            else if (wallDirectionType == "Y-WALL")
                            {
                                needsRotation = false;
                                DebugLogger.Warning($"[UniversalSleevePlacer] Fallback: Y-WALL detected - no rotation");
                        }
                        else
                        {
                                // Ultimate fallback
                                needsRotation = false;
                                DebugLogger.Warning($"[UniversalSleevePlacer] Unknown orientation - using fallback");
                            }
                        }
                        
                        // Log to placement_debug.log for immediate visibility
                        try
                        {
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
                                
                                DebugLogger.Info($"[ORIENT-FIX] Sleeve {sleeveInstance.Id.IntegerValue}: X-WALL +90° (LEFT view family)");
                                try
                                {
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
                            
                            DebugLogger.Info($"[ORIENT-SKIP] Sleeve {sleeveInstance.Id.IntegerValue}: Y-WALL no rotation (LEFT view family already aligned)");
                            try
                            {
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
                            DebugLogger.Info($"[UniversalSleevePlacer] WALL/FRAMING: Set HostOrientation = {clashZone.HostOrientation} (from XML)");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Warning($"[UniversalSleevePlacer] Error setting sleeve orientation: {ex.Message}\n{ex.StackTrace}");
                try
                {
                    DebugLogger.Info($"[ORIENT-ERROR] Sleeve {sleeveInstance.Id.IntegerValue}: {ex.Message}\n");
                }
                catch { }
            }
        }
    }
}
