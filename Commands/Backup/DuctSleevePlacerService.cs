using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Electrical;
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB.Structure;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using JSE_RevitAddin_MEP_OPENINGS.Services.ClearanceProviders;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Strategies;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// OPTIMIZED DuctSleevePlacerService - Uses pre-calculated data from ClashZone (no linked file access needed)
    /// </summary>
    public class DuctSleevePlacerService
    {
        private readonly Document _doc;
        private readonly OpeningConditions _conditions;
        public int PlacedCount { get; set; }
        public int SkippedCount { get; set; }
        public int ErrorCount { get; set; }

        public DuctSleevePlacerService(Document doc, OpeningConditions conditions = null)
        {
            DebugLogger.SetServiceContext("SleevePlacers");
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _conditions = conditions ?? new OpeningConditions(); // Use defaults if not provided
            DebugLogger.Info($"[DuctSleevePlacerService] OPTIMIZED Constructor called for document: {doc.Title}");
            DebugLogger.Info($"[DuctSleevePlacerService] Using clearances from CONDITIONS.xml - RectNormal:{_conditions.ClearanceSettings.RectangularNormal}mm, RoundNormal:{_conditions.ClearanceSettings.RoundNormal}mm");
        }

        /// <summary>
        /// Collect ducts that have corresponding clash zones from ALL documents (active + linked)
        /// </summary>
        public List<Duct> CollectDuctsWithClashZones(List<ClashZone> clashZones)
        {
            try
            {
                var ducts = new List<Duct>();
                var mepElementIds = clashZones.Select(cz => cz.MepElementId).ToHashSet();
                
                DebugLogger.Info($"[DuctSleevePlacerService] Looking for {mepElementIds.Count} ducts with IDs: {string.Join(", ", mepElementIds)}");
                
                // 1. Search in active document first
                var activeDucts = new FilteredElementCollector(_doc)
                    .OfClass(typeof(Duct))
                    .Cast<Duct>()
                    .Where(duct => mepElementIds.Contains(duct.Id))
                    .ToList();
                
                ducts.AddRange(activeDucts);
                DebugLogger.Info($"[DuctSleevePlacerService] Found {activeDucts.Count} ducts in active document");
                
                // 2. Search in linked documents
                var linkedDocs = _doc.Application.Documents.Cast<Document>()
                    .Where(doc => doc.IsLinked && doc.Title != _doc.Title)
                    .ToList();
                
                DebugLogger.Info($"[DuctSleevePlacerService] Searching {linkedDocs.Count} linked documents for ducts");
                
                foreach (var linkedDoc in linkedDocs)
                {
                    try
                    {
                        var linkedDucts = new FilteredElementCollector(linkedDoc)
                            .OfClass(typeof(Duct))
                            .Cast<Duct>()
                            .Where(duct => mepElementIds.Contains(duct.Id))
                                    .ToList();
                        
                        ducts.AddRange(linkedDucts);
                        DebugLogger.Info($"[DuctSleevePlacerService] Found {linkedDucts.Count} ducts in linked document: {linkedDoc.Title}");
                                }
                                catch (Exception ex)
                                {
                        DebugLogger.Warning($"[DuctSleevePlacerService] Error searching ducts in linked document {linkedDoc.Title}: {ex.Message}");
                    }
                }

                DebugLogger.Info($"[DuctSleevePlacerService] Total collected {ducts.Count} ducts with clash zones from all documents");
                return ducts;
                        }
                        catch (Exception ex)
                        {
                DebugLogger.Error($"[DuctSleevePlacerService] Error collecting ducts with clash zones: {ex.Message}");
                return new List<Duct>();
            }
        }

        /// <summary>
        /// OPTIMIZED: Place all sleeves using pre-calculated data (no linked file access needed)
        /// </summary>
        public int PlaceAllSleevesOptimized(List<Duct> ducts, Transaction transaction, List<ClashZone> clashZones = null, string ductOpeningType = "Circular")
        {
            if (transaction == null)
                throw new ArgumentNullException(nameof(transaction));

            if (!transaction.HasStarted())
                throw new InvalidOperationException("Transaction must be started before calling this method");

            // Use provided clash zones
            var availableClashZones = clashZones ?? new List<ClashZone>();

            int placed = 0;
            int skipped = 0;
            int errors = 0;

            foreach (var inputDuct in ducts)
            {
                // Find corresponding clash zone for this duct (declare outside try block for catch access)
                var clashZone = availableClashZones?.FirstOrDefault(cz => cz.MepElementId == inputDuct.Id);
                try
                {
                    if (clashZone == null)
                    {
                        DebugLogger.Warning($"[DuctSleevePlacerServiceOptimized] No clash zone found for duct {inputDuct.Id}");
                        skipped++;
                    continue;
                }

                    // OPTIMIZATION: Check if already resolved (no duplicate detection needed)
                    if (clashZone.IsResolved)
                    {
                        DebugLogger.Info($"[DuctSleevePlacerServiceOptimized] SKIP: Clash zone {clashZone.Id} already resolved");
                        skipped++;
                            continue;
                        }
                        
                    // Get family symbols
                    var symbols = GetFamilySymbols();
                    if (symbols.wallSymbol == null || symbols.slabSymbol == null)
                        {
                        DebugLogger.Warning($"[DuctSleevePlacerServiceOptimized] Family symbols not available for duct {inputDuct.Id}");
                        skipped++;
                            continue;
                        }

                    // OPTIMIZATION: Use pre-calculated data (NO LINKED FILE ACCESS NEEDED)
                    var placementPoint = clashZone.SleevePlacementPoint;
                    var rawWidth = clashZone.MepElementWidth;
                    var rawHeight = clashZone.MepElementHeight;
                    var mepOrientation = clashZone.MepElementOrientation;
                    var pipeOpeningType = clashZone.PipeOpeningType;

                    // ⚠️ CRITICAL: Validate clash zone category matches this service ⚠️
                    // DO NOT REMOVE: Prevents cross-category contamination (e.g., pipes/cable trays in duct service)
                    // Each service should only process its own category (Ducts, Pipes, Cable Trays, etc.)
                    if (!string.IsNullOrEmpty(clashZone.MepElementCategory) && 
                        clashZone.MepElementCategory != "Ducts")
                    {
                        DebugLogger.Error($"[DuctSleevePlacerServiceOptimized] ClashZone {clashZone.Id} is for category '{clashZone.MepElementCategory}', not 'Ducts' - SKIPPING");
                        skipped++;
                        continue;
                    }

                    // Use the duct from the input list (already collected from linked documents)
                    // Note: inputDuct is guaranteed to be a Duct because this service only handles ducts
                    // The ductsWithClashZones list is pre-filtered by DuctSleevePlacementCommand
                    var mepElement = inputDuct as Element;

                    // Validate this is actually a Duct (defensive programming)
                    if (!(inputDuct is Duct))
                    {
                        DebugLogger.Error($"[DuctSleevePlacerServiceOptimized] inputDuct {inputDuct.Id} is not a Duct, it's a {inputDuct.GetType().Name} - SKIPPING");
                        errors++;
                        continue;
                    }
                    
                    DebugLogger.Info($"[DuctSleevePlacerService] Processing Duct {inputDuct.Id} from clash zone {clashZone.Id} (category: {clashZone.MepElementCategory})");

                    // ⚠️ CRITICAL OPTIMIZATION: Use pre-calculated dimensions from XML (already include clearance)
                    // This ensures consistency and avoids double-clearance calculation
                    var finalWidth = clashZone.MepElementWidth;   // Already includes clearance
                    var finalHeight = clashZone.MepElementHeight; // Already includes clearance
                    
                    // Debug: Log pre-calculated dimensions
                    double finalWidthMm = UnitUtils.ConvertFromInternalUnits(finalWidth, UnitTypeId.Millimeters);
                    double finalHeightMm = UnitUtils.ConvertFromInternalUnits(finalHeight, UnitTypeId.Millimeters);
                    DebugLogger.Info($"[DuctSleevePlacerService] Using pre-calculated dimensions from XML: Width={finalWidthMm:F0}mm, Height={finalHeightMm:F0}mm (already includes clearance)");

                    // ⚠️ CRITICAL: Determine opening type based on duct shape from family name ⚠️
                    // DO NOT REMOVE: This logic ensures correct sleeve family selection
                    // Round ducts → UI selection, Rectangular ducts → always rectangular
                    string openingType = "Rectangular"; // Default for ducts
                    if (mepElement is Pipe)
                    {
                        openingType = pipeOpeningType; // Use pipe opening type from UI
                    }
                    else if (mepElement is Duct ductElement)
                    {
                        // Get duct shape from ClashZone (determined from family name during refresh)
                        var ductShape = clashZone.DuctShape ?? "Rectangular";
                        
                        DebugLogger.Info($"[DuctSleevePlacerService] Duct {mepElement.Id} shape from family name: {ductShape}");
                        DebugLogger.Info($"[DuctSleevePlacerService] UI duct opening type selection: {ductOpeningType}");
                        
                        // For Round Ducts: Check UI selection (Circular or Rectangular)
                        if (ductShape == "Round")
                        {
                            openingType = ductOpeningType; // Use UI selection for round ducts
                            DebugLogger.Info($"[DuctSleevePlacerService] Round duct -> {openingType} opening (from UI selection)");
                        }
                        else // Rectangular Ducts
                        {
                            openingType = "Rectangular"; // Rectangular ducts always use rectangular openings
                            DebugLogger.Info($"[DuctSleevePlacerService] Rectangular duct -> Rectangular opening (forced)");
                        }
                    }

                    // Validate mepElement is not null before proceeding
                    if (mepElement == null)
                    {
                        DebugLogger.Error($"[DuctSleevePlacerServiceOptimized] mepElement is NULL for clash zone {clashZone.Id} - element ID {clashZone.MepElementId} not found in document");
                        errors++;
                        continue;
                    }

                    // Get appropriate family symbol using element types
                    DebugLogger.Info($"[DuctSleevePlacerService] ClashZone data check - Id: {clashZone.Id}, StructuralType: {clashZone.StructuralElementType ?? "NULL"}, OpeningType: {openingType}");
                    DebugLogger.Info($"[DuctSleevePlacerService] About to get symbol for: MEP={mepElement.Id}, Structural={clashZone.StructuralElementType}, OpeningType={openingType}");
                    FamilySymbol appropriateSymbol = GetAppropriateSymbol(mepElement, clashZone.StructuralElementType, openingType, symbols);
                    if (appropriateSymbol == null)
                    {
                        DebugLogger.Warning($"[DuctSleevePlacerServiceOptimized] No appropriate symbol found for MEP element {mepElement.Id} and structural element type {clashZone.StructuralElementType} with opening type {openingType}");
                        skipped++;
                            continue;
                        }
                    
                    DebugLogger.Info($"[DuctSleevePlacerService] Selected symbol: {appropriateSymbol.FamilyName} for opening type {openingType}");

                    // Validate pre-calculated data before placement
                    if (placementPoint == null)
                    {
                        DebugLogger.Error($"[DuctSleevePlacerServiceOptimized] placementPoint is null for clash zone {clashZone.Id}");
                        errors++;
                        continue;
                    }
                    if (mepOrientation == null)
                    {
                        DebugLogger.Error($"[DuctSleevePlacerServiceOptimized] mepOrientation is null for clash zone {clashZone.Id}");
                        errors++;
                        continue;
                    }

                    DebugLogger.Info($"[DuctSleevePlacerService] Validated data - PlacementPoint: {placementPoint}, Orientation: {mepOrientation}, Level: {clashZone.MepElementLevelName ?? "null"}");

                    // Validate mepElement is actually a Duct
                    var duct = mepElement as Duct;
                    if (duct == null)
                    {
                        DebugLogger.Error($"[DuctSleevePlacerServiceOptimized] mepElement {mepElement.Id} is not a Duct, it's a {mepElement.GetType().Name}");
                        errors++;
                        continue;
                    }

                    // OPTIMIZATION: Use UniversalSleevePlacerService for "calculate once, use many times" principle
                    var universalService = new UniversalSleevePlacerService(_doc, _conditions, new DuctPlacementStrategy());
                    var result = universalService.PlaceAllSleevesInTransaction(new List<ClashZone> { clashZone });
                    placed += result.PlacedCount;

                    // Mark as resolved after successful placement
                    clashZone.IsResolved = true;

                    DebugLogger.Info($"[DuctSleevePlacerServiceOptimized] OPTIMIZED: Placed sleeve for element {mepElement.Id} at {placementPoint}");
                    }
                    catch (Exception ex)
                    {
                    errors++;
                    var elementId = clashZone.MepElementId.ToString();
                    DebugLogger.Error($"[DuctSleevePlacerServiceOptimized] Failed to place sleeve for element {elementId}: {ex.Message}");
                    DebugLogger.Error($"[DuctSleevePlacerServiceOptimized] Stack trace: {ex.StackTrace}");
                }
            }
            
            DebugLogger.Info($"[DuctSleevePlacerServiceOptimized] OPTIMIZED Placement complete: {placed} placed, {skipped} skipped, {errors} errors");
            return placed;
        }

        /// <summary>
        /// Get appropriate family symbol using element types and opening type (for pipes/ducts)
        /// </summary>
        private FamilySymbol GetAppropriateSymbol(Element mepElement, string structuralElementType, string openingType, (FamilySymbol? wallSymbol, FamilySymbol? slabSymbol) symbols)
        {
            try
            {
                // 1. Use the passed structural element type (Wall/Floor/Framing/None)
                string structuralType = structuralElementType;
                
                // 2. Check if openings are allowed on this structural element
                if (structuralType == "Unknown")
                {
                    DebugLogger.Warning($"[DuctSleevePlacerServiceOptimized] No openings allowed on structural element type {structuralElementType} (likely a column)");
                    return null; // No family symbol - skip placement
                }
                
                // 3. Determine MEP element type (Duct/Pipe/CableTray)
                string mepType = GetMepElementType(mepElement);
                
                // 4. Build family name
                string familyName = BuildFamilyName(mepType, structuralType, openingType);
                
                // 5. Get family symbol by name
                return GetFamilySymbolByName(familyName, symbols);
                }
                catch (Exception ex)
                {
                DebugLogger.Error($"[DuctSleevePlacerServiceOptimized] Error determining appropriate symbol: {ex.Message}");
                return symbols.wallSymbol; // Default fallback
            }
        }


        /// <summary>
        /// Check if structural framing is a column
        /// </summary>
        private bool IsColumn(FamilyInstance framing)
        {
            // Check if the structural framing is vertical (column)
            if (framing.Location is LocationCurve curve)
            {
                var direction = curve.Curve.GetEndPoint(1) - curve.Curve.GetEndPoint(0);
                return Math.Abs(direction.Z) > 0.9; // Nearly vertical (column)
            }
            
            // Alternative: Check family name for column indicators
            var familyName = framing.Symbol.FamilyName.ToLower();
            return familyName.Contains("column") || familyName.Contains("post");
        }

        /// <summary>
        /// Get MEP element type
        /// </summary>
        private string GetMepElementType(Element mepElement)
        {
            if (mepElement is Duct) return "Duct";
            if (mepElement is Pipe) return "Pipe";
            if (mepElement is Conduit) return "CableTray";
            return "Duct"; // Default
        }

        /// <summary>
        /// Build family name based on MEP type, structural type, and opening type
        /// </summary>
        private string BuildFamilyName(string mepType, string structuralType, string openingType)
        {
            string baseName = $"{mepType}OpeningOn{structuralType}";
            
            DebugLogger.Info($"[DuctSleevePlacerService] Building family name: MEP={mepType}, Structural={structuralType}, Opening={openingType}");
            
            // Special cases for rectangular/circular openings
            if (mepType == "Pipe" && openingType == "Rectangular")
            {
                baseName = $"PipeOpeningOn{structuralType}Rectangular"; // e.g., PipeOpeningOnWallRectangular
                DebugLogger.Info($"[DuctSleevePlacerService] Pipe rectangular -> {baseName}");
            }
            else if (mepType == "Duct" && openingType == "Rectangular")
            {
                // For rectangular ducts - use standard rectangular family
                baseName = $"DuctOpeningOn{structuralType}"; // e.g., DuctOpeningOnWall (standard rectangular)
                DebugLogger.Info($"[DuctSleevePlacerService] Duct rectangular -> {baseName}");
            }
            else if (mepType == "Duct" && openingType == "Circular")
            {
                // For round ducts (<200mm) when user selects circular
                baseName = $"DuctOpeningOn{structuralType}round"; // e.g., DuctOpeningOnWallround (lowercase to match project)
                DebugLogger.Info($"[DuctSleevePlacerService] Duct circular -> {baseName}");
            }
            else
            {
                DebugLogger.Info($"[DuctSleevePlacerService] Default case -> {baseName}");
            }
            
            DebugLogger.Info($"[DuctSleevePlacerService] Final family name: {baseName}");
            return baseName;
        }

        /// <summary>
        /// Get family symbol by name
        /// </summary>
        private FamilySymbol GetFamilySymbolByName(string familyName, (FamilySymbol? wallSymbol, FamilySymbol? slabSymbol) symbols)
        {
            try
            {
                DebugLogger.Info($"[DuctSleevePlacerService] Searching for family symbol: {familyName}");
                
                // Search for the specific family symbol by name (case-insensitive)
                var familySymbol = new FilteredElementCollector(_doc)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>()
                    .FirstOrDefault(fs => fs.FamilyName.Equals(familyName, StringComparison.OrdinalIgnoreCase));

                if (familySymbol != null)
                {
                    DebugLogger.Info($"[DuctSleevePlacerService] Found exact family symbol: {familyName}");
                    return familySymbol;
                }

                // Fallback to simple logic based on structural type
                DebugLogger.Warning($"[DuctSleevePlacerService] Family symbol '{familyName}' not found, using fallback");
                DebugLogger.Info($"[DuctSleevePlacerService] Available wall symbol: {symbols.wallSymbol?.FamilyName ?? "null"}");
                DebugLogger.Info($"[DuctSleevePlacerService] Available slab symbol: {symbols.slabSymbol?.FamilyName ?? "null"}");
                
                if (familyName.Contains("OnWall"))
                {
                    DebugLogger.Info($"[DuctSleevePlacerService] Using wall symbol fallback: {symbols.wallSymbol?.FamilyName ?? "null"}");
                    return symbols.wallSymbol;
                }
                else if (familyName.Contains("OnSlab"))
                {
                    DebugLogger.Info($"[DuctSleevePlacerService] Using slab symbol fallback: {symbols.slabSymbol?.FamilyName ?? "null"}");
                    return symbols.slabSymbol;
                }
                
                DebugLogger.Info($"[DuctSleevePlacerService] Using default wall symbol fallback: {symbols.wallSymbol?.FamilyName ?? "null"}");
                return symbols.wallSymbol; // Default fallback
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[DuctSleevePlacerService] Error finding family symbol '{familyName}': {ex.Message}");
                return symbols.wallSymbol; // Default fallback
            }
        }

        /// <summary>
        /// Get family symbols for duct sleeves based on opening type (circular/rectangular)
        /// </summary>
        private (FamilySymbol? wallSymbol, FamilySymbol? slabSymbol) GetFamilySymbols(string openingType = "Circular")
        {
            try
            {
                FamilySymbol? wallSymbol = null;
                FamilySymbol? slabSymbol = null;

                if (openingType == "Circular")
                {
                    // For circular ducts, prefer DuctOpeningOnWallround and DuctOpeningOnSlabround (case-insensitive)
                    wallSymbol = new FilteredElementCollector(_doc)
                        .OfClass(typeof(FamilySymbol))
                        .Cast<FamilySymbol>()
                        .FirstOrDefault(fs => fs.FamilyName.Contains("DuctOpeningOnWallround", StringComparison.OrdinalIgnoreCase));

                    slabSymbol = new FilteredElementCollector(_doc)
                        .OfClass(typeof(FamilySymbol))
                        .Cast<FamilySymbol>()
                        .FirstOrDefault(fs => fs.FamilyName.Contains("DuctOpeningOnSlabround", StringComparison.OrdinalIgnoreCase));

                    // Fallback to generic circular families if specific ones not found
                    if (wallSymbol == null)
                    {
                        wallSymbol = new FilteredElementCollector(_doc)
                            .OfClass(typeof(FamilySymbol))
                            .Cast<FamilySymbol>()
                            .FirstOrDefault(fs => fs.FamilyName.Contains("DuctOpeningOnWall") && !fs.FamilyName.Contains("Rect"));
                    }

                    if (slabSymbol == null)
                    {
                        slabSymbol = new FilteredElementCollector(_doc)
                            .OfClass(typeof(FamilySymbol))
                            .Cast<FamilySymbol>()
                            .FirstOrDefault(fs => fs.FamilyName.Contains("DuctOpeningOnSlab") && !fs.FamilyName.Contains("Rect"));
                    }
                }
                else // Rectangular
                {
                    // For rectangular ducts, prefer DuctOpeningOnWallRect and DuctOpeningOnSlabRect
                    wallSymbol = new FilteredElementCollector(_doc)
                        .OfClass(typeof(FamilySymbol))
                        .Cast<FamilySymbol>()
                        .FirstOrDefault(fs => fs.FamilyName.Contains("DuctOpeningOnWallRect"));

                    slabSymbol = new FilteredElementCollector(_doc)
                        .OfClass(typeof(FamilySymbol))
                        .Cast<FamilySymbol>()
                        .FirstOrDefault(fs => fs.FamilyName.Contains("DuctOpeningOnSlabRect"));

                    // Fallback to generic rectangular families if specific ones not found
                    if (wallSymbol == null)
                    {
                        wallSymbol = new FilteredElementCollector(_doc)
                            .OfClass(typeof(FamilySymbol))
                            .Cast<FamilySymbol>()
                            .FirstOrDefault(fs => fs.FamilyName.Contains("DuctOpeningOnWall") && !fs.FamilyName.Contains("Round"));
                    }

                    if (slabSymbol == null)
                    {
                        slabSymbol = new FilteredElementCollector(_doc)
                            .OfClass(typeof(FamilySymbol))
                            .Cast<FamilySymbol>()
                            .FirstOrDefault(fs => fs.FamilyName.Contains("DuctOpeningOnSlab") && !fs.FamilyName.Contains("Round"));
                    }
                }

                DebugLogger.Info($"[DuctSleevePlacerService] Found family symbols for {openingType} ducts - Wall: {wallSymbol?.FamilyName ?? "null"}, Slab: {slabSymbol?.FamilyName ?? "null"}");
                
                // DEBUG: List all available duct opening family symbols
                var allDuctFamilies = new FilteredElementCollector(_doc)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>()
                    .Where(fs => fs.FamilyName.Contains("DuctOpening"))
                    .Select(fs => fs.FamilyName)
                    .ToList();
                
                DebugLogger.Info($"[DuctSleevePlacerService] Available duct opening families: {string.Join(", ", allDuctFamilies)}");
                
                return (wallSymbol, slabSymbol);
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[DuctSleevePlacerService] Error getting family symbols: {ex.Message}");
                return (null, null);
            }
        }

        /// <summary>
        /// ⚠️ CRITICAL METHOD - DO NOT REMOVE ⚠️
        /// Get clearance from CONDITIONS.xml based on duct shape and insulation type
        /// Implements proper architecture: Conditions saved in XML, not static properties
        /// </summary>
        private double GetClearanceFromUI(string ductShape, string insulationType)
        {
            try
            {
                double clearanceInMm = 50.0; // Default
                
                // Select appropriate clearance from CONDITIONS.xml based on duct shape and insulation
                if (ductShape == "Round")
                {
                    clearanceInMm = insulationType == "Insulated" 
                        ? _conditions.ClearanceSettings.RoundInsulated 
                        : _conditions.ClearanceSettings.RoundNormal;
                }
                else // Rectangular
                {
                    clearanceInMm = insulationType == "Insulated" 
                        ? _conditions.ClearanceSettings.RectangularInsulated 
                        : _conditions.ClearanceSettings.RectangularNormal;
                }
                
                DebugLogger.Info($"[GetClearanceFromUI] Using clearance from CONDITIONS.xml: {clearanceInMm}mm (shape: {ductShape}, insulation: {insulationType})");
                
                // Convert from mm to feet (Revit internal units)
                return UnitUtils.ConvertToInternalUnits(clearanceInMm, UnitTypeId.Millimeters);
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[GetClearanceFromUI] Error: {ex.Message}, using default 50mm");
                return UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters);
            }
        }
    }
}
