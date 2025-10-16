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
        
        public int PlacedCount { get; private set; }
        public int SkippedCount { get; private set; }
        public int ErrorCount { get; private set; }

        public UniversalSleevePlacerService(Document doc, OpeningConditions conditions, ISleevePlacementStrategy strategy)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _conditions = conditions ?? new OpeningConditions();
            _strategy = strategy ?? throw new ArgumentNullException(nameof(strategy));
            
            DebugLogger.Info($"[UniversalSleevePlaycer] Initialized for category: {_strategy.GetCategoryName()}");
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
            
            try
            {
                // ⚠️ CRITICAL: Iterate over CLASH ZONES, not MEP elements
                // Each clash zone represents a unique (MEP Element + Structural Element) PAIR
                // The same MEP element can appear in multiple clash zones if it intersects multiple walls
                // Clash zones are already filtered by UI during refresh, so process all provided zones
                foreach (var clashZone in clashZones)
                {
                    try
                    {
                        DebugLogger.Info($"[UniversalSleevePlacer] Processing ClashZone {clashZone.Id}: MEP={clashZone.MepElementId.IntegerValue}, Structural={clashZone.StructuralElementId.IntegerValue}");
                        
                        // ⚠️ CRITICAL: Hierarchical flag check - cluster takes precedence
                        File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log", 
                            $"[HIER-CHECK] ClashZone {clashZone.Id}: IsClusterResolved={clashZone.IsClusterResolved}, ClusterSleeveInstanceId={clashZone.ClusterSleeveInstanceId}\n");
                        
                        // STEP 1: Check if this clash zone is part of a cluster
                        if (clashZone.IsClusterResolved && clashZone.ClusterSleeveInstanceId > 0)
                        {
                            // Check if cluster sleeve still exists in Revit
                            var clusterSleeveId = new ElementId(clashZone.ClusterSleeveInstanceId);
                            var clusterSleeve = _doc.GetElement(clusterSleeveId);
                            
                            File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log", 
                                $"[HIER-CHECK] ClashZone {clashZone.Id}: Checking cluster sleeve {clashZone.ClusterSleeveInstanceId}, exists={clusterSleeve != null}\n");
                            
                            if (clusterSleeve != null)
                            {
                                // Cluster sleeve exists - SKIP this clash zone (cluster handles it)
                                DebugLogger.Info($"[UniversalSleevePlacer] SKIP: ClashZone {clashZone.Id} is part of cluster sleeve {clashZone.ClusterSleeveInstanceId}");
                                File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log", 
                                    $"✓ SKIP ClashZone {clashZone.Id}: cluster sleeve {clashZone.ClusterSleeveInstanceId} exists\n");
                                SkippedCount++;
                                continue;
                            }
                            else
                            {
                                // Cluster sleeve was deleted - reset BOTH flags and place individual sleeve
                                DebugLogger.Info($"[UniversalSleevePlacer] Cluster sleeve {clashZone.ClusterSleeveInstanceId} was deleted - resetting flags and placing individual sleeve");
                                clashZone.IsClusterResolved = false;
                                clashZone.ClusterSleeveInstanceId = -1;
                                clashZone.IsResolved = false; // ⚠️ CRITICAL: Also reset individual flag
                                clashZone.SleeveInstanceId = -1;
                                // Continue to place individual sleeve below
                            }
                        }
                        
                        // STEP 2: Check if individual sleeve already exists (only if NOT cluster-resolved)
                        if (clashZone.IsResolved && clashZone.SleeveInstanceId > 0)
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
                                // Individual sleeve was deleted - reset flag and place new one
                                DebugLogger.Info($"[UniversalSleevePlacer] Individual sleeve {clashZone.SleeveInstanceId} was deleted - resetting flag and placing new sleeve");
                                clashZone.IsResolved = false;
                                clashZone.SleeveInstanceId = -1;
                                // Continue to place individual sleeve below
                            }
                        }
                        
                        // STEP 3: If we reach here, place individual sleeve (fresh or replacement)
                        
                        // Validate category match
                        if (!string.IsNullOrEmpty(clashZone.MepElementCategory) && 
                            clashZone.MepElementCategory != _strategy.GetCategoryName())
                        {
                            DebugLogger.Warning($"[UniversalSleevePlacer] SKIP: ClashZone {clashZone.Id} category '{clashZone.MepElementCategory}' doesn't match '{_strategy.GetCategoryName()}'");
                            SkippedCount++;
                            continue;
                        }
                        
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
                        double finalWidth, finalHeight, finalDiameter;
                        
                        // 🛡️ ARCHITECTURE FIX: Use CONDITIONS service for ALL clearance types
                        // This ensures consistent architecture: CONDITIONS XML → UniversalSleevePlacerService
                        // Raw dimensions from ClashZone + Clearance from CONDITIONS = Final dimensions
                        
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
                        else if (_strategy is CableTrayPlacementStrategy cableTrayStrategy)
                        {
                            // ✅ Cable trays: Raw dimensions + CONDITIONS clearance via strategy
                            var rawWidth = clashZone.MepElementWidth;
                            var rawHeight = clashZone.MepElementHeight;
                            
                            // Get offset and final dimensions from strategy (uses CONDITIONS)
                            var adj2 = cableTrayStrategy.GetCableTrayPlacementAdjustment(clashZone, _conditions);
                            placementOffset = adj2.offsetVector;
                            finalWidth = adj2.finalWidth;
                            finalHeight = adj2.finalHeight;
                            finalDiameter = finalWidth; // Not used for cable trays (rectangular only)
                            
                            DebugLogger.Info($"[UniversalSleevePlacer] CABLE TRAY: Raw={UnitUtils.ConvertFromInternalUnits(rawWidth, UnitTypeId.Millimeters):F1}x{UnitUtils.ConvertFromInternalUnits(rawHeight, UnitTypeId.Millimeters):F1}mm → Final={UnitUtils.ConvertFromInternalUnits(finalWidth, UnitTypeId.Millimeters):F1}x{UnitUtils.ConvertFromInternalUnits(finalHeight, UnitTypeId.Millimeters):F1}mm");
                        }
                        else
                        {
                            // ✅ Ducts: Raw dimensions + CONDITIONS clearance
                            var rawWidth = clashZone.MepElementWidth;
                            var rawHeight = clashZone.MepElementHeight;
                            
                            DebugLogger.Info($"[UniversalSleevePlacer] DUCT CALCULATION START: Raw dimensions {UnitUtils.ConvertFromInternalUnits(rawWidth, UnitTypeId.Millimeters):F1}x{UnitUtils.ConvertFromInternalUnits(rawHeight, UnitTypeId.Millimeters):F1}mm");
                            
                            var clearance = GetClearanceFromConditions("Ducts", mepSize);
                            finalWidth = rawWidth + (2 * clearance);
                            finalHeight = rawHeight + (2 * clearance);
                            finalDiameter = finalWidth; // For round elements
                            
                            DebugLogger.Info($"[UniversalSleevePlacer] DUCT: Raw={UnitUtils.ConvertFromInternalUnits(rawWidth, UnitTypeId.Millimeters):F1}x{UnitUtils.ConvertFromInternalUnits(rawHeight, UnitTypeId.Millimeters):F1}mm + Clearance={UnitUtils.ConvertFromInternalUnits(clearance, UnitTypeId.Millimeters):F1}mm = Final={UnitUtils.ConvertFromInternalUnits(finalWidth, UnitTypeId.Millimeters):F1}x{UnitUtils.ConvertFromInternalUnits(finalHeight, UnitTypeId.Millimeters):F1}mm");
                        }

                        // ⚠️ REMOVED: Old width/height swapping logic that was causing double-swapping
                        // The new logic later in the method (lines 660-665) handles this correctly
                        // by ensuring the longer dimension becomes width, not just swapping blindly
                        
                        // Select universal family
                        var (familyName, typeName) = SelectUniversalFamily(clashZone, mepSize);
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
                        SetSleeveParameters(sleeveInstance, mepSize, finalWidth, finalHeight, finalDiameter, clashZone);
                        
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
        
        /// <summary>
        /// Select universal family based on host type and MEP shape
        /// Uses 4 universal families following CONVOID approach
        /// </summary>
        private (string familyName, string typeName) SelectUniversalFamily(ClashZone clashZone, MepElementSize mepSize)
        {
            // Determine host type (check both singular and plural forms)
            bool isWallOrFraming = clashZone.StructuralElementType == "Wall" || 
                                  clashZone.StructuralElementType == "Walls" ||
                                  clashZone.StructuralElementType == "Structural Framing";
            
            // 🛡️ ARCHITECTURE FIX: Use CONDITIONS XML for opening type preferences (not ClashZone or geometry)
            // This follows the reference architecture: CONDITIONS.xml stores user preferences
            bool isCircular;
            if (string.Equals(clashZone.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase))
            {
                // ✅ CORRECT: Pipes opening type from CONDITIONS XML (user preference)
                var pipeType = _conditions?.OpeningTypePreferences?.Pipes ?? "Circular";
                isCircular = string.Equals(pipeType, "Circular", StringComparison.OrdinalIgnoreCase);
                DebugLogger.Info($"[UniversalSleevePlacer] PIPE opening type from CONDITIONS XML: '{pipeType}' → isCircular={isCircular}");
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
            
            return (familyName, typeName);
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
        /// Get clearance value from CONDITIONS service for simple clearance categories
        /// </summary>
        private double GetClearanceFromConditions(string category, MepElementSize mepSize)
        {
            try
            {
                DebugLogger.Info($"[GetClearanceFromConditions] START: category='{category}', _conditions={(_conditions != null ? "NOT NULL" : "NULL")}");
                
                if (_conditions?.ClearanceSettings == null)
                {
                    DebugLogger.Warning($"[GetClearanceFromConditions] No clearance settings available, using default 50mm");
                    return UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters);
                }

                DebugLogger.Info($"[GetClearanceFromConditions] ClearanceSettings available: RectNormal={_conditions.ClearanceSettings.RectangularNormal}mm, RectInsulated={_conditions.ClearanceSettings.RectangularInsulated}mm");

                // Determine clearance based on category and element properties
                double clearanceInMm = 50.0; // Default fallback

                if (string.Equals(category, "Pipes", StringComparison.OrdinalIgnoreCase))
                {
                    // Pipes: Check if insulated
                    bool isInsulated = IsPipeInsulated(mepSize);
                    clearanceInMm = isInsulated ? _conditions.ClearanceSettings.PipesInsulated : _conditions.ClearanceSettings.PipesNormal;
                    DebugLogger.Info($"[GetClearanceFromConditions] Pipes: isInsulated={isInsulated}, clearance={clearanceInMm}mm");
                }
                else if (string.Equals(category, "Ducts", StringComparison.OrdinalIgnoreCase))
                {
                    // Ducts: Check if insulated and shape
                    bool isInsulated = IsDuctInsulated(mepSize);
                    bool isRound = string.Equals(mepSize.Shape, "Round", StringComparison.OrdinalIgnoreCase) ||
                                  string.Equals(mepSize.Shape, "Circular", StringComparison.OrdinalIgnoreCase);
                    
                    DebugLogger.Info($"[GetClearanceFromConditions] Ducts: isInsulated={isInsulated}, isRound={isRound}, Shape='{mepSize.Shape}'");
                    
                    if (isRound)
                    {
                        clearanceInMm = isInsulated ? _conditions.ClearanceSettings.RoundInsulated : _conditions.ClearanceSettings.RoundNormal;
                        DebugLogger.Info($"[GetClearanceFromConditions] Round ducts: clearance={clearanceInMm}mm");
                    }
                    else
                    {
                        clearanceInMm = isInsulated ? _conditions.ClearanceSettings.RectangularInsulated : _conditions.ClearanceSettings.RectangularNormal;
                        DebugLogger.Info($"[GetClearanceFromConditions] Rectangular ducts: clearance={clearanceInMm}mm");
                    }
                }

                // Convert from mm to feet (Revit internal units)
                double clearanceInFeet = UnitUtils.ConvertToInternalUnits(clearanceInMm, UnitTypeId.Millimeters);
                
                DebugLogger.Info($"[GetClearanceFromConditions] {category}: {clearanceInMm}mm → {clearanceInFeet:F6}ft");
                return clearanceInFeet;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[GetClearanceFromConditions] Error getting clearance for {category}: {ex.Message}");
                return UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters); // Safe fallback
            }
        }

        /// <summary>
        /// Determine if a pipe is insulated (simplified logic)
        /// </summary>
        private bool IsPipeInsulated(MepElementSize mepSize)
        {
            // Simplified logic - in real implementation, this would check pipe parameters
            // For now, assume larger pipes are more likely to be insulated
            double diameterMm = UnitUtils.ConvertFromInternalUnits(mepSize.Diameter, UnitTypeId.Millimeters);
            
            // Assume pipes > 200mm are insulated
            return diameterMm > 200.0;
        }

        /// <summary>
        /// Determine if a duct is insulated (simplified logic)
        /// </summary>
        private bool IsDuctInsulated(MepElementSize mepSize)
        {
            // Simplified logic - in real implementation, this would check duct parameters
            // For now, assume larger ducts are more likely to be insulated
            double maxDimension = Math.Max(mepSize.Width, mepSize.Height);
            double maxDimensionMm = UnitUtils.ConvertFromInternalUnits(maxDimension, UnitTypeId.Millimeters);
            
            // Assume ducts > 500mm are insulated
            return maxDimensionMm > 500.0;
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
            ClashZone clashZone)
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
            isPipe;

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
                    $"[DEPTH-SET] Sleeve {sleeveInstance.Id.IntegerValue}: {(isFramingHost ? "FRAMING" : "FLOOR")} - Set Depth = {depthReadMm:F1}mm, Verified={verified} {(verified ? "✓" : "✗")}\n");
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
            
                    if (mepOrientation != null && (mepOrientation.X != 0 || mepOrientation.Y != 0))
            {
                var loc = sleeveInstance.Location as LocationPoint;
                if (loc != null)
                {
                    // SPECIAL CASE: Pipes and Cable Trays on Framing - align to MEP direction
                    if ((isPipe || isCableTray) && isFramingHost)
                    {
                        double angle = Math.Atan2(mepOrientation.Y, mepOrientation.X);
                        
                        // For framing, we need to consider the structural normal to determine correct rotation
                        if (structuralNormal != null)
                        {
                            // Calculate the angle between MEP direction and structural normal
                            double mepAngle = Math.Atan2(mepOrientation.Y, mepOrientation.X);
                            double structAngle = Math.Atan2(structuralNormal.Y, structuralNormal.X);
                            
                            // The sleeve should be perpendicular to the framing direction
                            // If MEP is parallel to framing, rotate 90°
                            double dotProduct = mepOrientation.X * structuralNormal.X + mepOrientation.Y * structuralNormal.Y;
                            if (Math.Abs(dotProduct) < 0.1) // Nearly perpendicular (MEP ⊥ Framing)
                            {
                                angle = mepAngle; // Use MEP direction as-is
                            }
                            else // Nearly parallel (MEP ∥ Framing)
                            {
                                angle = mepAngle + Math.PI / 2; // Rotate 90° from MEP direction
                            }
                        }
                        
                        double angleDegrees = angle * 180 / Math.PI;
                        
                        Line rotationAxis = Line.CreateBound(loc.Point, loc.Point + XYZ.BasisZ);
                        ElementTransformUtils.RotateElement(_doc, sleeveInstance.Id, rotationAxis, angle);
                        
                        string mepType = isPipe ? "PIPE" : "CABLETRAY";
                        DebugLogger.Info($"[UniversalSleevePlacer] FRAMING+{mepType}: Aligned sleeve to MEP direction {angleDegrees:F1}°");
                        try
                        {
                            System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log",
                                $"[ORIENT-APPLY] Sleeve {sleeveInstance.Id.IntegerValue}: FRAMING+{mepType} aligned to MEP {angleDegrees:F1}° ✓\n");
                        }
                        catch { }
                    }
                    else
                    {
                        // GENERAL CASE: Check if wall/framing is Y-oriented to decide rotation
                        bool allowRotate = false;
                        
                        if (structuralNormal != null && (structuralNormal.X != 0 || structuralNormal.Y != 0))
                        {
                            // FIX: Use wall direction logic instead of normal direction
                            // For X-wall: Direction=(1,0,0) or (-1,0,0), Normal=(0,1,0) or (0,-1,0)
                            // For Y-wall: Direction=(0,1,0) or (0,-1,0), Normal=(1,0,0) or (-1,0,0)
                            // We need to infer wall direction from normal: if normal is Y-oriented, wall is X-oriented
                            
                            double absWX = Math.Abs(structuralNormal.X);
                            double absWY = Math.Abs(structuralNormal.Y);
                            
                            // FIXED LOGIC: If normal is Y-oriented (|Wy| > |Wx|), then wall is X-oriented (no rotation)
                            // If normal is X-oriented (|Wx| > |Wy|), then wall is Y-oriented (rotation allowed)
                            allowRotate = absWX > absWY; // Rotate only for Y-oriented walls/framing
                            
                            // DEBUG: Explicit wall orientation detection
                            bool isXWall = absWY > absWX; // Normal is Y-oriented = Wall is X-oriented
                            bool isYWall = absWX > absWY; // Normal is X-oriented = Wall is Y-oriented
                            string wallType = isXWall ? "X-WALL" : (isYWall ? "Y-WALL" : "UNKNOWN");
                            
                            DebugLogger.Info($"[UniversalSleevePlacer] WALL/FRAMING GUARD: |Wx|={absWX:F3}, |Wy|={absWY:F3}, Type={wallType}, allowRotate={allowRotate}");
                            try
                            {
                                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log",
                                    $"[ORIENT-GUARD] Sleeve {sleeveInstance.Id.IntegerValue}: |Wx|={absWX:F3}, |Wy|={absWY:F3}, Type={wallType}, allowRotate={allowRotate}\n");
                            }
                            catch { }
                        }
                        
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
