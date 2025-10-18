using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Strategies;
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
        
        public int PlacedCount { get; private set; }
        public int SkippedCount { get; private set; }
        public int ErrorCount { get; private set; }

        public UniversalSleevePlacerService(Document doc, OpeningConditions conditions, ISleevePlacementStrategy strategy, Dictionary<string, double> clearanceSettings = null)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _conditions = conditions ?? new OpeningConditions();
            _strategy = strategy ?? throw new ArgumentNullException(nameof(strategy));
            _clearanceSettings = clearanceSettings ?? new Dictionary<string, double>();
            
            // 🔥 CRITICAL DEBUG: Direct file logging to trace service instantiation
            System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\service_instantiation.log", 
                $"[{DateTime.Now:HH:mm:ss}] 🔥 UniversalSleevePlacerService CONSTRUCTOR CALLED 🔥\n");
            System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\service_instantiation.log", 
                $"[{DateTime.Now:HH:mm:ss}] Strategy Type: {_strategy.GetType().Name}\n");
            System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\service_instantiation.log", 
                $"[{DateTime.Now:HH:mm:ss}] Strategy Category: {_strategy.GetCategoryName()}\n");
            System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\service_instantiation.log", 
                $"[{DateTime.Now:HH:mm:ss}] Clearance Settings Count: {_clearanceSettings.Count}\n");
            
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
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log",
                    $"[BUILD] {ts} Assembly={System.IO.Path.GetFileName(asm.Location)} Version={ver} Path={asm.Location}\n");
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
                    System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log",
                        $"[BUILD] {ts} Assembly={System.IO.Path.GetFileName(asm.Location)} Version={ver} Path={asm.Location}\n");
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
            PlacedCount = 0;
            SkippedCount = 0;
            ErrorCount = 0;
            
            // 🔥 CRITICAL DEBUG: Log flag status from the clash zones passed to this method
            int clusterResolvedCount = clashZones.Count(cz => cz.IsClusterResolved);
            int individualResolvedCount = clashZones.Count(cz => cz.IsResolved);
            File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log", 
                $"[{DateTime.Now:HH:mm:ss}] 🔥 UNIVERSAL SLEEVE PLACER RECEIVED: {clashZones.Count} clash zones\n");
            File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log", 
                $"[{DateTime.Now:HH:mm:ss}] 📊 FLAGS RECEIVED: IsClusterResolved=True: {clusterResolvedCount}, IsResolved=True: {individualResolvedCount}\n");
            
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
            
            try
            {
                // ⚠️ CRITICAL: Iterate over CLASH ZONES, not MEP elements
                // Each clash zone represents a unique (MEP Element + Structural Element) PAIR
                // The same MEP element can appear in multiple clash zones if it intersects multiple walls
                // Clash zones are already filtered by UI during refresh, so process all provided zones
                foreach (var clashZone in sortedClashZones)
                {
                    try
                    {
                        DebugLogger.Info($"[UniversalSleevePlacer] Processing ClashZone {clashZone.Id}: MEP={clashZone.MepElementId.IntegerValue}, Structural={clashZone.StructuralElementId.IntegerValue}");
                        
                        // ⚠️ CRITICAL: Hierarchical flag check - cluster takes precedence
                        File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log", 
                            $"[HIER-CHECK] ClashZone {clashZone.Id}: IsClustered={clashZone.IsClustered}, IsClusterResolved={clashZone.IsClusterResolved}, ClusterSleeveInstanceId={clashZone.ClusterSleeveInstanceId}\n");
                        
                        // STEP 0: Check if already clustered (prevent individual sleeves over cluster sleeves)
                        if (clashZone.IsClustered)
                        {
                            DebugLogger.Info($"[UniversalSleevePlacer] SKIP: ClashZone {clashZone.Id} is already clustered (ClusterSleeveInstanceId: {clashZone.ClusterSleeveInstanceId}) - preventing individual sleeve placement");
                            File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log", 
                                $"✓ SKIP ClashZone {clashZone.Id}: already clustered (ClusterSleeveInstanceId: {clashZone.ClusterSleeveInstanceId})\n");
                            SkippedCount++;
                            continue;
                        }
                        
                        // STEP 1: Check individual sleeve flag first (fastest check)
                        if (clashZone.IsResolved)
                        {
                            // ✅ CRITICAL: Check both flags according to flag management logic
                            if (clashZone.IsClusterResolved)
                            {
                                // Both flags TRUE: Skip processing (avoid clash zone)
                                File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log", 
                                    $"✓ SKIP ClashZone {clashZone.Id}: Both IsResolved=True AND IsClusterResolved=True - avoid clash zone\n");
                                SkippedCount++;
                                continue;
                            }
                            else
                            {
                                // Only individual flag TRUE: Skip individual sleeve placement
                                File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log", 
                                    $"✓ SKIP ClashZone {clashZone.Id}: IsResolved=True (individual sleeve already placed)\n");
                                SkippedCount++;
                                continue;
                            }
                        }
                        
                        // STEP 2: Check cluster sleeve flag
                        if (clashZone.IsClusterResolved)
                        {
                            // Only cluster flag TRUE: Skip both individual and cluster placement
                            File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log", 
                                $"✓ SKIP ClashZone {clashZone.Id}: IsClusterResolved=True (cluster sleeve handles it)\n");
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
                                    File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log", 
                                        $"✓ SKIP ClashZone {clashZone.Id}: Individual sleeve missing but cluster sleeve exists - keeping IsResolved=true\n");
                                    SkippedCount++;
                                    continue;
                                }
                            }
                        }
                        
                        // STEP 3: If we reach here, place individual sleeve (fresh or replacement)
                        
                        // 🔥 DEBUG: Log category comparison
                        System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\category_match_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ClashZone {clashZone.Id}: MepElementCategory='{clashZone.MepElementCategory}' (len={clashZone.MepElementCategory?.Length}), Strategy='{_strategy.GetCategoryName()}' (len={_strategy.GetCategoryName()?.Length})\n");
                        
                        // Validate category match
                        if (!string.IsNullOrEmpty(clashZone.MepElementCategory) && 
                            clashZone.MepElementCategory != _strategy.GetCategoryName())
                        {
                            System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\category_match_debug.log", 
                                $"[{DateTime.Now:HH:mm:ss}] 🔥 CATEGORY MISMATCH! Skipping ClashZone {clashZone.Id}\n");
                            DebugLogger.Warning($"[UniversalSleevePlacer] SKIP: ClashZone {clashZone.Id} category '{clashZone.MepElementCategory}' doesn't match '{_strategy.GetCategoryName()}'");
                            SkippedCount++;
                            continue;
                        }
                        
                        System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\category_match_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ✅ CATEGORY MATCH! Proceeding with ClashZone {clashZone.Id}\n");
                        
                        
                        // 🔥 DEBUG: Log that we're about to start clearance calculation
                        System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\clearance_calculation_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] 🔥 ABOUT TO START CLEARANCE CALCULATION for ClashZone {clashZone.Id}\n");
                        
                        // ⚠️ ZERO LINKED FILE ACCESS - use pre-calculated MEP size from ClashZone
                        var mepSize = new MepElementSize
                        {
                            Width = clashZone.MepElementWidth,
                            Height = clashZone.MepElementHeight,
                            Diameter = clashZone.MepElementWidth, // For round, width = diameter
                            Shape = clashZone.DuctShape
                        };
                        
						// ⚠️ SPECIAL HANDLING & SIZING ORDER:
						// 1) Pipes (host-agnostic), 2) Dampers, 3) Cable trays, 4) Ducts/default
                        XYZ placementOffset = XYZ.Zero;
                        double finalWidth = 0.0, finalHeight = 0.0, finalDiameter = 0.0;
                        
                        // 🛡️ ARCHITECTURE FIX: Use CONDITIONS service for ALL clearance types
                        // This ensures consistent architecture: CONDITIONS XML → UniversalSleevePlacerService
                        // Raw dimensions from ClashZone + Clearance from CONDITIONS = Final dimensions
                        
                        System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\clearance_calculation_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] 🔥 CLEARANCE CALCULATION START: Category='{clashZone.MepElementCategory}', Strategy={(_strategy?.GetType().Name ?? "NULL")}\n");
                        
                        DebugLogger.Info($"[UniversalSleevePlacer] CLEARANCE CALCULATION START: Category='{clashZone.MepElementCategory}', Strategy={(_strategy?.GetType().Name ?? "NULL")}");
                        
						bool isPipesCategory = string.Equals(clashZone.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase);
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
                            // 🔥 DEBUG: Log that we're entering the cable tray strategy block
                            System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\clearance_calculation_debug.log", 
                                $"[{DateTime.Now:HH:mm:ss}] 🎯 CABLE TRAY STRATEGY BLOCK ENTERED for ClashZone {clashZone.Id}\n");
                            System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\clearance_calculation_debug.log", 
                                $"[{DateTime.Now:HH:mm:ss}] About to call GetCableTrayPlacementAdjustment with {_clearanceSettings.Count} settings\n");
                            
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

                        // ⚠️ REMOVED: Old width/height swapping logic that was causing double-swapping
                        // The new logic later in the method (lines 660-665) handles this correctly
                        // by ensuring the longer dimension becomes width, not just swapping blindly
                        
                        // Select universal family
                        var (familyName, typeName, isCircular) = SelectUniversalFamily(clashZone, mepSize);
                        var familySymbol = LoadFamilySymbol(familyName);
                        
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
                        
                        // Find nearest level
                        var nearestLevel = FindNearestLevel(adjustedPlacementPoint);
                        if (nearestLevel == null)
                        {
                            DebugLogger.Warning($"[UniversalSleevePlacer] No level found for placement point");
                            SkippedCount++;
                            continue;
                        }
                        
                        // Place sleeve instance (NO HOST PARAMETER - workplane-based families)
                        // ✅ Works with linked structural elements because no host reference needed
                        var sleeveInstance = _doc.Create.NewFamilyInstance(
                            adjustedPlacementPoint,
                            familySymbol,
                            nearestLevel,
                            StructuralType.NonStructural);
                        
                        if (sleeveInstance == null)
                        {
                            DebugLogger.Error($"[UniversalSleevePlacer] Failed to create sleeve instance");
                            ErrorCount++;
                            continue;
                        }
                        
                        // Set parameters
                        SetSleeveParameters(sleeveInstance, mepSize, finalWidth, finalHeight, finalDiameter, clashZone, isCircular);
                        
                        // ⚠️ CRITICAL: Set orientation (rotation for floors, HostOrientation parameter for walls/framing)
                        SetSleeveOrientation(sleeveInstance, clashZone);

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
                            }
                        }
                        catch { }
                        
                        // Update ClashZone flags
                        clashZone.IsResolved = true;
                        clashZone.SleeveInstanceId = sleeveInstance.Id.IntegerValue;
                        clashZone.SleeveFamilyName = familySymbol.Family.Name;
                        
                        PlacedCount++;
                        DebugLogger.Info($"[UniversalSleevePlacer] ✓ Placed {_strategy.GetCategoryName()} sleeve {sleeveInstance.Id} for ClashZone {clashZone.Id} at {adjustedPlacementPoint}");
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Error($"[UniversalSleevePlacer] Error placing sleeve for ClashZone {clashZone.Id}: {ex.Message}");
                        ErrorCount++;
                    }
                }
                
                DebugLogger.Info($"[UniversalSleevePlacer] Placement loop complete - Placed: {PlacedCount}, Skipped: {SkippedCount}, Errors: {ErrorCount}");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UniversalSleevePlacer] Error in placement loop: {ex.Message}");
                throw;
            }
            
            return (PlacedCount, SkippedCount);
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
            var filtersDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings", "Projects", "Default", "Filters");
            
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
        /// Check for existing cluster sleeves by spatial proximity (handles stale XML flags)
        /// This ensures individual sleeves are not placed over cluster sleeves even when XML is stale
        /// </summary>
        private bool CheckForClusterSleeveByProximity(ClashZone clashZone)
        {
            try
            {
                // Get placement point
                var placementPoint = clashZone.SleevePlacementPoint;
                if (placementPoint == null) return false;
                
                // Search for cluster sleeves within 500mm radius
                double searchRadius = UnitUtils.ConvertToInternalUnits(500.0, UnitTypeId.Millimeters);
                
                // Create bounding box for spatial search
                var searchBox = new BoundingBoxXYZ
                {
                    Min = new XYZ(placementPoint.X - searchRadius, placementPoint.Y - searchRadius, placementPoint.Z - searchRadius),
                    Max = new XYZ(placementPoint.X + searchRadius, placementPoint.Y + searchRadius, placementPoint.Z + searchRadius)
                };
                
                // Collect cluster sleeve families
                var clusterFamilies = new[] { "DuctOpeningOnWall", "DuctOpeningOnFloor", "PipeOpeningOnWallRect", "PipeOpeningOnFloorRect", "CableTrayOpeningOnWall", "CableTrayOpeningOnFloor" };
                
                var collector = new FilteredElementCollector(_doc)
                    .WherePasses(new BoundingBoxIntersectsFilter(new Outline(searchBox.Min, searchBox.Max)))
                    .OfClass(typeof(FamilyInstance));
                
                foreach (FamilyInstance instance in collector)
                {
                    if (instance.Symbol?.Family?.Name != null)
                    {
                        string familyName = instance.Symbol.Family.Name;
                        
                        // Check if this is a cluster sleeve family
                        if (clusterFamilies.Any(cf => familyName.Contains(cf)))
                        {
                            // Check if it's on the same host element
                            var hostElement = instance.Host;
                            if (hostElement != null && hostElement.Id == clashZone.StructuralElementId)
                            {
                                // Found a cluster sleeve on the same host - skip individual sleeve placement
                                File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log", 
                                    $"[SPATIAL-CHECK] ClashZone {clashZone.Id}: Found cluster sleeve {instance.Id} ({familyName}) on same host {hostElement.Id}\n");
                                return true;
                            }
                        }
                    }
                }
                
                return false;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UniversalSleevePlacer] Error checking for cluster sleeve by proximity: {ex.Message}");
                return false;
            }
        }

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
        
        private FamilySymbol LoadFamilySymbol(string familyName)
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
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\clearance_calculation.log", 
                    $"[{DateTime.Now:HH:mm:ss}] GetClearanceFromConditions START: category='{category}', Shape='{mepSize.Shape}', IsInsulated={mepSize.IsInsulated}\n");
                
                DebugLogger.Info($"[GetClearanceFromConditions] START: category='{category}', UI settings={_clearanceSettings.Count}, XML conditions={(_conditions != null ? "NOT NULL" : "NULL")}");
                
                // ✅ PRIORITY 1: Use UI clearance settings if available
                if (_clearanceSettings.Count > 0)
                {
                    double clearanceInMm = GetClearanceFromUISettings(category, mepSize);
                    if (clearanceInMm > 0)
                    {
                        System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\clearance_calculation.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ✅ Using UI clearance: {clearanceInMm}mm for {category}\n");
                        DebugLogger.Info($"[GetClearanceFromConditions] Using UI clearance: {clearanceInMm}mm for {category}");
                        return UnitUtils.ConvertToInternalUnits(clearanceInMm, UnitTypeId.Millimeters);
                    }
                }
                
                // ✅ PRIORITY 2: Fallback to XML conditions
                if (_conditions?.ClearanceSettings != null)
                {
                    System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\clearance_calculation.log", 
                        $"[{DateTime.Now:HH:mm:ss}] Using XML clearance settings for {category}\n");
                    DebugLogger.Info($"[GetClearanceFromConditions] Using XML clearance settings for {category}");
                    return GetClearanceFromXmlConditions(category, mepSize);
                }
                
                // ✅ PRIORITY 3: Default fallback
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\clearance_calculation.log", 
                    $"[{DateTime.Now:HH:mm:ss}] ⚠️ No clearance settings available, using default 50mm\n");
                DebugLogger.Warning($"[GetClearanceFromConditions] No clearance settings available, using default 50mm");
                return UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters);
            }
            catch (Exception ex)
            {
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\clearance_calculation.log", 
                    $"[{DateTime.Now:HH:mm:ss}] ❌ ERROR in clearance calculation: {ex.Message}\n");
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
            System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\clearance_calculation.log", 
                $"[{DateTime.Now:HH:mm:ss}] IsPipeInsulated: Using actual data - IsInsulated={mepSize.IsInsulated}, InsulationThickness={mepSize.InsulationThickness:F6}ft\n");
            
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
            System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\clearance_calculation.log", 
                $"[{DateTime.Now:HH:mm:ss}] IsDuctInsulated: Using actual data - IsInsulated={mepSize.IsInsulated}, InsulationThickness={mepSize.InsulationThickness:F6}ft\n");
            
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
            System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log",
                $"[PARAM-SET] Sleeve {sleeveInstance.Id.IntegerValue}: isPipe={isPipe}, Shape={mepSize.Shape}, roundedDia={diamMm:F1}mm, W={widthMm:F1}mm, H={heightMm:F1}mm\n");
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
                        System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log",
                            $"[PARAM-SUCCESS] Sleeve {sleeveInstance.Id.IntegerValue}: Set '{name}' = {diamMm:F1}mm\n");
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
        
        double thicknessMm = UnitUtils.ConvertFromInternalUnits(clashZone.StructuralElementThickness, UnitTypeId.Millimeters);
        
        try
        {
            System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log",
                $"[DEPTH-CHECK] Sleeve {sleeveInstance.Id.IntegerValue}: Host={clashZone.StructuralElementType}, isWall={isWallHost}, isFraming={isFramingHost}, HasDepth={depthParam != null}, DepthRO={depthParam?.IsReadOnly}, HasWallWidth={wallWidthParam != null}, WallWidthRO={wallWidthParam?.IsReadOnly}, StructThickness={thicknessMm:F1}mm\n");
        }
        catch { }
        
        bool depthSetSuccess = false;
        
        if (isWallHost && wallWidthParam != null && !wallWidthParam.IsReadOnly)
        {
            // Wall host: use Wall Width parameter
                    wallWidthParam.Set(clashZone.StructuralElementThickness);
            DebugLogger.Info($"[UniversalSleevePlacer] WALL: Set Wall Width = {thicknessMm:F1}mm");
            try
            {
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log",
                    $"[DEPTH-SET] Sleeve {sleeveInstance.Id.IntegerValue}: WALL - Set Wall Width = {thicknessMm:F1}mm ✓\n");
            }
            catch { }
            depthSetSuccess = true;
                }
                else if (depthParam != null && !depthParam.IsReadOnly)
                {
            // Floor/Framing host: use Depth parameter
                    depthParam.Set(clashZone.StructuralElementThickness);
            DebugLogger.Info($"[UniversalSleevePlacer] {(isFramingHost ? "FRAMING" : "FLOOR")}: Set Depth = {thicknessMm:F1}mm");
            try
            {
                // Verify the parameter was set correctly
                double depthRead = depthParam.AsDouble();
                double depthReadMm = UnitUtils.ConvertFromInternalUnits(depthRead, UnitTypeId.Millimeters);
                bool verified = Math.Abs(depthRead - clashZone.StructuralElementThickness) < 0.0001;
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
                typeDepthParam.Set(clashZone.StructuralElementThickness);
                DebugLogger.Info($"[UniversalSleevePlacer] Set TYPE Depth = {thicknessMm:F1}mm (instance param not writable)");
                try
                {
                    _doc.Regenerate();
                    System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log",
                        $"[DEPTH-SET] Sleeve {sleeveInstance.Id.IntegerValue}: TYPE Depth = {thicknessMm:F1}mm (regenerated) ✓\n");
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
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log",
                    $"[DEPTH-FAIL] Sleeve {sleeveInstance.Id.IntegerValue}: No writable depth parameter found! ✗\n");
            }
            catch { }
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
                                    System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log",
                                        $"[HOST-PARAM-SET] Sleeve {sleeveInstance.Id.IntegerValue}: Set '{hostParam.Key}' = '{hostParam.Value}' ({param.StorageType}) ✓\n");
                                }
                                catch { }
                            }
                            else
                            {
                                DebugLogger.Warning($"[UniversalSleevePlacer] Host parameter '{hostParam.Key}' not found or read-only on sleeve {sleeveInstance.Id}");
                                try
                                {
                                    System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log",
                                        $"[HOST-PARAM-MISSING] Sleeve {sleeveInstance.Id.IntegerValue}: Parameter '{hostParam.Key}' not found or read-only ✗\n");
                                }
                                catch { }
                            }
                        }
                        catch (Exception paramEx)
                        {
                            DebugLogger.Warning($"[UniversalSleevePlacer] Error setting host parameter '{hostParam.Key}' = '{hostParam.Value}': {paramEx.Message}");
                            try
                            {
                                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log",
                                    $"[HOST-PARAM-ERROR] Sleeve {sleeveInstance.Id.IntegerValue}: '{hostParam.Key}' = '{hostParam.Value}' - {paramEx.Message} ✗\n");
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
                        System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log",
                            $"[HOST-PARAM-EMPTY] Sleeve {sleeveInstance.Id.IntegerValue}: No HostParameterValues in ClashZone {clashZone.Id} ✗\n");
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
            System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log",
                $"[PARAM-ERROR] Sleeve {sleeveInstance.Id.IntegerValue}: {ex.Message}\n");
        }
        catch { }
    }
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
                    System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log",
                        $"[FLOOR-DEBUG] Sleeve {sleeveInstance.Id.IntegerValue}: StructuralElementType='{clashZone.StructuralElementType}', isFloorHost={isFloorHost}\n");
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
                            
                            // ⚠️ SIMPLIFIED LOGIC FOR FLOORS: Focus on vertical elements only
                            // All vertical elements (ducts, pipes, cable trays) use the same logic
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
                            
                            double rotationAngleDegrees = rotationAngle * 180 / Math.PI;
                            
                            Line rotationAxis = Line.CreateBound(loc.Point, loc.Point + XYZ.BasisZ);
                            ElementTransformUtils.RotateElement(_doc, sleeveInstance.Id, rotationAxis, rotationAngle);
                            
                    DebugLogger.Info($"[UniversalSleevePlacer] FLOOR: Rotated sleeve {rotationAngleDegrees:F1}° based on MEP orientation ({mepOrientation})");
                    try
                    {
                        System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log",
                            $"[ORIENT-FLOOR] Sleeve {sleeveInstance.Id.IntegerValue}: Rotated {rotationAngleDegrees:F1}° for MEP dir ({mepOrientation})\n");
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
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log",
                    $"[WALL-DEBUG] Sleeve {sleeveInstance.Id.IntegerValue}: StructuralElementType='{clashZone.StructuralElementType}', isFramingHost={isFramingHost}\n");
            }
            catch { }
            bool isPipe = string.Equals(clashZone.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase);
            bool isCableTray = string.Equals(clashZone.MepElementCategory, "Cable Trays", StringComparison.OrdinalIgnoreCase) ||
                              string.Equals(clashZone.MepElementCategory, "Cable Tray Fittings", StringComparison.OrdinalIgnoreCase);
                    
                    var mepOrientation = clashZone.MepElementOrientation;
            var structuralNormal = clashZone.StructuralElementNormal;
            
            try
            {
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log",
                    $"[ORIENT-INPUT] Sleeve {sleeveInstance.Id.IntegerValue}: Cat={clashZone.MepElementCategory}, Host={clashZone.StructuralElementType}, MEP=({mepOrientation?.X:F3},{mepOrientation?.Y:F3}), StructN=({structuralNormal?.X:F3},{structuralNormal?.Y:F3})\n");
                
                // DEBUG: Log all ClashZone properties to trace XML loading issue
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log",
                    $"[XML-DEBUG] ClashZone {clashZone.Id}: MepElementCategory='{clashZone.MepElementCategory}', StructuralElementType='{clashZone.StructuralElementType}', MepElementId={clashZone.MepElementIdValue}, StructuralElementId={clashZone.StructuralElementIdValue}\n");
            }
            catch { }
            
                    if (true) // DISABLED: Don't wait for MEP orientation - use wall direction directly
            {
                var loc = sleeveInstance.Location as LocationPoint;
                if (loc != null)
                {
                    // ✅ OPTIMIZED: Pipes and Cable Trays on Framing - use pre-calculated framing direction from XML
                    if ((isPipe || isCableTray) && isFramingHost)
                    {
                        // Use pre-calculated framing direction from XML (no Revit API calls!)
                        XYZ framingDirection = clashZone.StructuralElementNormal; // StructuralElementNormal stores framing direction for framing hosts
                        
                        if (framingDirection != null && framingDirection != XYZ.Zero)
                        {
                            // Calculate angle from pre-calculated framing direction
                            double angle = Math.Atan2(framingDirection.Y, framingDirection.X);
                            
                            // ✅ CRITICAL FIX: For pipes on structural framing in X direction, add 90° rotation
                            if (isPipe)
                            {
                                // Check if framing is in X direction (horizontal)
                                double absX = Math.Abs(framingDirection.X);
                                double absY = Math.Abs(framingDirection.Y);
                                
                                if (absX > absY)
                                {
                                    // X-direction framing: add 90° rotation for pipes
                                    angle += Math.PI / 2; // Add 90 degrees
                                    DebugLogger.Info($"[UniversalSleevePlacer] FRAMING+PIPE: X-direction framing detected - adding 90° rotation");
                                }
                            }
                            
                            double angleDegrees = angle * 180 / Math.PI;
                            
                            Line rotationAxis = Line.CreateBound(loc.Point, loc.Point + XYZ.BasisZ);
                            ElementTransformUtils.RotateElement(_doc, sleeveInstance.Id, rotationAxis, angle);
                            
                            string mepType = isPipe ? "PIPE" : "CABLETRAY";
                            DebugLogger.Info($"[UniversalSleevePlacer] FRAMING+{mepType}: Using pre-calculated framing direction {angleDegrees:F1}° from XML");
                            try
                            {
                                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log",
                                    $"[ORIENT-APPLY] Sleeve {sleeveInstance.Id.IntegerValue}: FRAMING+{mepType} using XML framing direction {angleDegrees:F1}° ✓\n");
                            }
                            catch { }
                        }
                        else
                        {
                            // Fallback: Use MEP direction if framing direction not available in XML
                            double angle = Math.Atan2(mepOrientation.Y, mepOrientation.X);
                            
                            // ✅ CRITICAL FIX: For pipes on structural framing, add 90° rotation in fallback too
                            if (isPipe)
                            {
                                // For pipes, add 90° rotation as default behavior for structural framing
                                angle += Math.PI / 2; // Add 90 degrees
                                DebugLogger.Info($"[UniversalSleevePlacer] FRAMING+PIPE: Fallback mode - adding 90° rotation for pipes");
                            }
                            
                            double angleDegrees = angle * 180 / Math.PI;
                            
                            Line rotationAxis = Line.CreateBound(loc.Point, loc.Point + XYZ.BasisZ);
                            ElementTransformUtils.RotateElement(_doc, sleeveInstance.Id, rotationAxis, angle);
                            
                            string mepType = isPipe ? "PIPE" : "CABLETRAY";
                            DebugLogger.Warning($"[UniversalSleevePlacer] FRAMING+{mepType}: Framing direction not in XML, using MEP direction {angleDegrees:F1}° (fallback)");
                        }
                    }
                    else
                    {
                        // DEBUG: Confirm we're reaching the wall direction logic
                        try
                        {
                            System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log",
                                $"[WALL-LOGIC-REACHED] Sleeve {sleeveInstance.Id.IntegerValue}: Reached wall direction logic block ✓\n");
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
                                    System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log",
                                        $"[WALL-DIR-FALLBACK] Sleeve {sleeveInstance.Id.IntegerValue}: Calculated WallDirection=({wallDirection.X:F3},{wallDirection.Y:F3},{wallDirection.Z:F3}), WallType={wallDirectionType} ✓\n");
                                }
                                catch { }
                            }
                            else
                            {
                                // Ultimate fallback
                                wallDirectionType = "UNKNOWN";
                                try
                                {
                                    System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log",
                                        $"[WALL-DIR-ERROR] Sleeve {sleeveInstance.Id.IntegerValue}: Cannot determine wall direction - structural normal is null or zero ✗\n");
                                }
                                catch { }
                            }
                        }
                        
                        // Use calculated wall direction type (from XML or fallback)
                        string wallType = wallDirectionType;
                        
                        DebugLogger.Info($"[UniversalSleevePlacer] WALL/FRAMING: Using wall direction type: {wallType}");
                        
                        // Log to placement_debug.log for immediate visibility
                            try
                            {
                                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log",
                                $"[WALL-DIR-FINAL] Sleeve {sleeveInstance.Id.IntegerValue}: WallDirection=({wallDirection?.X:F3},{wallDirection?.Y:F3},{wallDirection?.Z:F3}), WallType={wallType} ✓\n");
                            }
                            catch { }
                        
                        // Robust wall orientation logic based on wall direction type
                        if (wallType == "X-WALL")
                        {
                            allowRotate = false; // X-walls need +90° rotation
                            DebugLogger.Info($"[UniversalSleevePlacer] X-WALL detected: Will apply +90° rotation");
                        }
                        else if (wallType == "Y-WALL")
                        {
                            allowRotate = true; // Y-walls work naturally with LEFT view families
                            DebugLogger.Info($"[UniversalSleevePlacer] Y-WALL detected: No rotation needed");
                        }
                        else if (wallType == "FRAMING")
                        {
                            // For framing, use MEP orientation logic (like old DuctSleevePlacerService)
                            var framingMepOrientation = clashZone.MepElementOrientation;
                            if (framingMepOrientation != null)
                            {
                                bool isYAxisMep = Math.Abs(framingMepOrientation.Y) > Math.Abs(framingMepOrientation.X);
                                allowRotate = isYAxisMep; // Rotate only for Y-axis MEP elements
                                DebugLogger.Info($"[UniversalSleevePlacer] FRAMING: MEP orientation Y-axis={isYAxisMep}, allowRotate={allowRotate}");
                            }
                        }
                        else
                        {
                            // Fallback to old logic for unknown types
                            allowRotate = false;
                            DebugLogger.Warning($"[UniversalSleevePlacer] Unknown wall type '{wallType}', using fallback logic");
                        }
                        
                        // Log to placement_debug.log for immediate visibility
                        try
                        {
                            System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log",
                                $"[ORIENT-DECISION] Sleeve {sleeveInstance.Id.IntegerValue}: WallType={wallType}, allowRotate={allowRotate} ✓\n");
                        }
                        catch { }
                        
                        if (!allowRotate)
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
                                    System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log",
                                        $"[ORIENT-FIX] Sleeve {sleeveInstance.Id.IntegerValue}: X-WALL +90° (LEFT view family) ✓\n");
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
                                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log",
                                    $"[ORIENT-SKIP] Sleeve {sleeveInstance.Id.IntegerValue}: Y-WALL no rotation (LEFT view family) ✓\n");
                            }
                            catch { }
                        }
                    }
                }
            }
            
            // Set HostOrientation parameter based on structural normal
                    if (structuralNormal != null && (structuralNormal.X != 0 || structuralNormal.Y != 0))
                    {
                        double absX = Math.Abs(structuralNormal.X);
                        double absY = Math.Abs(structuralNormal.Y);
                        string orientation = absX > absY ? "X" : "Y";
                        
                        var hostOrientationParam = sleeveInstance.LookupParameter("HostOrientation");
                        if (hostOrientationParam != null && !hostOrientationParam.IsReadOnly)
                        {
                            hostOrientationParam.Set(orientation);
                            DebugLogger.Info($"[UniversalSleevePlacer] WALL/FRAMING: Set HostOrientation = {orientation}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Warning($"[UniversalSleevePlacer] Error setting sleeve orientation: {ex.Message}\n{ex.StackTrace}");
                try
                {
                    System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log",
                        $"[ORIENT-ERROR] Sleeve {sleeveInstance.Id.IntegerValue}: {ex.Message}\n");
                }
                catch { }
            }
        }
    }
}

