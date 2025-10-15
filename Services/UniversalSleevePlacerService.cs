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
            
            if (clashZones == null || clashZones.Count == 0)
            {
                DebugLogger.Warning($"[UniversalSleevePlacer] No clash zones provided");
                return (0, 0);
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
                        
						bool isPipesCategory = string.Equals(clashZone.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase);
						if (isPipesCategory)
						{
							var clearance = _strategy.GetClearance(mepSize, _conditions);
							var baseOd = GetPipeOutsideDiameterFromClashZone(clashZone);
							var ins = mepSize.IsInsulated ? mepSize.InsulationThickness : 0.0;
							finalDiameter = baseOd + (2 * ins) + (2 * clearance);
							finalWidth = finalDiameter;
							finalHeight = finalDiameter;
							try
							{
								System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log",
									$"[PIPE-SIZE] CZ={clashZone.Id} OD={baseOd:F6}ft, ins={ins:F6}ft, clr={clearance:F6}ft, finalDia={finalDiameter:F6}ft\n");
							}
							catch { }
						}
						else if (_strategy is DamperPlacementStrategy damperStrategy)
                        {
                            // Fire dampers: asymmetric clearance + offset for MSFD
                            var adj = damperStrategy.GetDamperPlacementAdjustment(clashZone, _conditions);
                            placementOffset = adj.offsetVector;
                            finalWidth = adj.finalWidth;
                            finalHeight = adj.finalHeight;
                            finalDiameter = adj.finalWidth; // Not used for dampers (rectangular only)
                        }
                        else if (_strategy is CableTrayPlacementStrategy cableTrayStrategy)
                        {
                            // Cable trays: asymmetric clearance + upward offset
                            var adj2 = cableTrayStrategy.GetCableTrayPlacementAdjustment(clashZone, _conditions);
                            placementOffset = adj2.offsetVector;
                            finalWidth = adj2.finalWidth;
                            finalHeight = adj2.finalHeight;
                            finalDiameter = adj2.finalWidth; // Not used for cable trays (rectangular only)
                        }
                        else
                        {
							// Ducts/default: symmetric clearance, no offset
                        var clearance = _strategy.GetClearance(mepSize, _conditions);
                            finalWidth = mepSize.Width + (2 * clearance);
                            finalHeight = mepSize.Height + (2 * clearance);
                            finalDiameter = mepSize.Diameter + (2 * clearance);
                        }

                        // OPTIONAL ORIENTATION NORMALIZATION FOR FLOOR DUCTS (rectangular):
                        // Swap width/height so the family's "Width" axis aligns with the MEP direction
                        // This avoids the need for a +90° correction for rectangular duct sleeves on floors.
                        try
                        {
                            bool isFloorHostHere = clashZone.StructuralElementType == "Floor" ||
                                                   clashZone.StructuralElementType == "Floors";
                            bool isDuctStrategy = _strategy is Strategies.DuctPlacementStrategy;
                            bool isRectangular = !(string.Equals(mepSize.Shape, "Round", StringComparison.OrdinalIgnoreCase) ||
                                                   string.Equals(mepSize.Shape, "Circular", StringComparison.OrdinalIgnoreCase));

                            if (isFloorHostHere && isDuctStrategy && isRectangular)
                            {
                                var tmp = finalWidth; finalWidth = finalHeight; finalHeight = tmp;
                                DebugLogger.Info("[UniversalSleevePlacer] FLOOR/DUCT (rectangular): swapped Width/Height to align family axis with MEP direction");
                            }
                        }
                        catch { }
                        
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
            
            // Determine opening shape per category (separate concerns for pipes vs ducts)
            bool isCircular;
            if (string.Equals(clashZone.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase))
            {
                // Pipes: driven by UI/global setting only
                var pipeType = clashZone.PipeOpeningType;
                if (string.IsNullOrWhiteSpace(pipeType))
                {
                    pipeType = OpeningSettingsHelper.GetOpeningTypeForCategory("Pipes");
                }
                isCircular = string.Equals(pipeType, "Circular", StringComparison.OrdinalIgnoreCase);
            }
            else if (string.Equals(clashZone.MepElementCategory, "Ducts", StringComparison.OrdinalIgnoreCase))
            {
                // Ducts: geometry-driven (round vs rectangular); UI may override elsewhere
                isCircular = string.Equals(mepSize.Shape, "Round", StringComparison.OrdinalIgnoreCase) ||
                             string.Equals(mepSize.Shape, "Circular", StringComparison.OrdinalIgnoreCase);
            }
            else
            {
                // Other categories (cable trays, accessories): rectangular
                isCircular = false;
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
        
        // CRITICAL: Set Depth parameter based on host type
        // Priority: Wall Width (walls) > Depth (floors/framing) > Type Depth (fallback)
        bool isWallHost = clashZone.StructuralElementType == "Wall" || clashZone.StructuralElementType == "Walls";
        bool isFramingHost = string.Equals(clashZone.StructuralElementType, "Structural Framing", StringComparison.OrdinalIgnoreCase);
        
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
            // ====== FLOOR HOST: Rotate based on MEP direction ======
                    var mepOrientation = clashZone.MepElementOrientation;
                    if (mepOrientation != null && (mepOrientation.X != 0 || mepOrientation.Y != 0))
                    {
                        var loc = sleeveInstance.Location as LocationPoint;
                        if (loc != null)
                        {
                            double rotationAngle = Math.Atan2(mepOrientation.Y, mepOrientation.X);
                            double rotationAngleDegrees = rotationAngle * 180 / Math.PI;
                            
                            Line rotationAxis = Line.CreateBound(loc.Point, loc.Point + XYZ.BasisZ);
                            ElementTransformUtils.RotateElement(_doc, sleeveInstance.Id, rotationAxis, rotationAngle);
                            
                    DebugLogger.Info($"[UniversalSleevePlacer] FLOOR: Rotated sleeve {rotationAngleDegrees:F1}° based on MEP orientation ({mepOrientation.X:F3}, {mepOrientation.Y:F3})");
                    try
                    {
                        System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log",
                            $"[ORIENT-FLOOR] Sleeve {sleeveInstance.Id.IntegerValue}: Rotated {rotationAngleDegrees:F1}° for MEP dir ({mepOrientation.X:F3},{mepOrientation.Y:F3})\n");
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
            }
            catch { }
            
                    if (mepOrientation != null && (mepOrientation.X != 0 || mepOrientation.Y != 0))
            {
                var loc = sleeveInstance.Location as LocationPoint;
                if (loc != null)
                {
                    // SPECIAL CASE: Pipes on Framing - always align to MEP direction
                    if (isPipe && isFramingHost)
                    {
                        double angle = Math.Atan2(mepOrientation.Y, mepOrientation.X);
                        double angleDegrees = angle * 180 / Math.PI;
                        
                        Line rotationAxis = Line.CreateBound(loc.Point, loc.Point + XYZ.BasisZ);
                        ElementTransformUtils.RotateElement(_doc, sleeveInstance.Id, rotationAxis, angle);
                        
                        DebugLogger.Info($"[UniversalSleevePlacer] FRAMING+PIPE: Aligned sleeve to MEP direction {angleDegrees:F1}°");
                        try
                        {
                            System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\placement_debug.log",
                                $"[ORIENT-APPLY] Sleeve {sleeveInstance.Id.IntegerValue}: FRAMING+PIPE aligned to MEP {angleDegrees:F1}° ✓\n");
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
