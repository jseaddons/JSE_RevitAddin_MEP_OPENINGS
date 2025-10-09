using System;
using System.Collections.Generic;
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
                foreach (var clashZone in clashZones)
                {
                    try
                    {
                        DebugLogger.Info($"[UniversalSleevePlacer] Processing ClashZone {clashZone.Id}: MEP={clashZone.MepElementId.IntegerValue}, Structural={clashZone.StructuralElementId.IntegerValue}");
                        
                        // ⚠️ LAYER 1: Trust flags completely (1,000,000x faster than Layer 2)
                        if (clashZone.IsResolved || clashZone.IsClustered)
                        {
                            DebugLogger.Info($"[UniversalSleevePlacer] SKIP: ClashZone {clashZone.Id} already resolved or clustered (IsResolved={clashZone.IsResolved}, IsClustered={clashZone.IsClustered}, IsClusterResolved={clashZone.IsClusterResolved})");
                            SkippedCount++;
                            continue;
                        }
                        
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
                        
                        // ⚠️ SPECIAL HANDLING: Dampers and Cable Trays use asymmetric clearance and need placement offset
                        XYZ placementOffset = XYZ.Zero;
                        double finalWidth, finalHeight, finalDiameter;
                        
                        if (_strategy is DamperPlacementStrategy damperStrategy)
                        {
                            // Fire dampers: asymmetric clearance + offset for MSFD
                            var (offsetVector, width, height) = damperStrategy.GetDamperPlacementAdjustment(clashZone, _conditions);
                            placementOffset = offsetVector;
                            finalWidth = width;
                            finalHeight = height;
                            finalDiameter = width; // Not used for dampers (rectangular only)
                        }
                        else if (_strategy is CableTrayPlacementStrategy cableTrayStrategy)
                        {
                            // Cable trays: asymmetric clearance + upward offset
                            var (offsetVector, width, height) = cableTrayStrategy.GetCableTrayPlacementAdjustment(clashZone, _conditions);
                            placementOffset = offsetVector;
                            finalWidth = width;
                            finalHeight = height;
                            finalDiameter = width; // Not used for cable trays (rectangular only)
                        }
                        else
                        {
                            // Ducts/Pipes: symmetric clearance, no offset
                        var clearance = _strategy.GetClearance(mepSize, _conditions);
                            finalWidth = mepSize.Width + (2 * clearance);
                            finalHeight = mepSize.Height + (2 * clearance);
                            finalDiameter = mepSize.Diameter + (2 * clearance);
                        }
                        
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
                        
                        // Apply placement offset for dampers/cable trays
                        XYZ adjustedPlacementPoint = clashZone.SleevePlacementPoint + placementOffset;
                        
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
            
            // Determine shape
            bool isCircular = mepSize.Shape == "Round" || mepSize.Shape == "Circular";
            
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
                // ⚠️ Apply rounding to nearest 5mm if setting is enabled
                var (roundedWidth, roundedHeight) = OpeningSettingsHelper.RoundDimensionsToNearest5mm(finalWidth, finalHeight);
                var roundedDiameter = OpeningSettingsHelper.RoundDiameterToNearest5mm(finalDiameter);
                
                // Set size parameters
                if (mepSize.Shape == "Round" || mepSize.Shape == "Circular")
                {
                    // Try Diameter parameter first (for circular openings)
                    var diamParam = sleeveInstance.LookupParameter("Diameter");
                    if (diamParam != null && !diamParam.IsReadOnly)
                    {
                        diamParam.Set(roundedDiameter);
                    }
                    else
                    {
                        // Fallback to Width/Height for circular
                        sleeveInstance.LookupParameter("Width")?.Set(roundedDiameter);
                        sleeveInstance.LookupParameter("Height")?.Set(roundedDiameter);
                    }
                }
                else
                {
                    // Rectangular
                    sleeveInstance.LookupParameter("Width")?.Set(roundedWidth);
                    sleeveInstance.LookupParameter("Height")?.Set(roundedHeight);
                }
                
                // ⚠️ CRITICAL: Set Depth parameter to structural element thickness
                // For walls: set to "Wall Width" parameter (wall thickness)
                // For floors/slabs: set to "Depth" parameter (floor thickness)
                // For framing: set to "b" parameter (beam width)
                var depthParam = sleeveInstance.LookupParameter("Depth");
                var wallWidthParam = sleeveInstance.LookupParameter("Wall Width");
                var bParam = sleeveInstance.LookupParameter("b");
                
                if (wallWidthParam != null && !wallWidthParam.IsReadOnly)
                {
                    wallWidthParam.Set(clashZone.StructuralElementThickness);
                    DebugLogger.Info($"[UniversalSleevePlacer] Set Wall Width = {clashZone.StructuralElementThickness} ft");
                }
                else if (depthParam != null && !depthParam.IsReadOnly)
                {
                    depthParam.Set(clashZone.StructuralElementThickness);
                    DebugLogger.Info($"[UniversalSleevePlacer] Set Depth = {clashZone.StructuralElementThickness} ft");
                }
                else if (bParam != null && !bParam.IsReadOnly)
                {
                    bParam.Set(clashZone.StructuralElementThickness);
                    DebugLogger.Info($"[UniversalSleevePlacer] Set b = {clashZone.StructuralElementThickness} ft");
                }
                
                // ⚠️ ZERO LINKED FILE ACCESS - all data pre-calculated during refresh!
                var mepElementIdParam = sleeveInstance.LookupParameter("MEP_ElementId");
                if (mepElementIdParam != null && !mepElementIdParam.IsReadOnly)
                {
                    mepElementIdParam.Set(clashZone.MepElementId.IntegerValue);
                    DebugLogger.Info($"[UniversalSleevePlacer] Set MEP_ElementId = {clashZone.MepElementId.IntegerValue} on sleeve {sleeveInstance.Id}");
                }
                else
                {
                    DebugLogger.Warning($"[UniversalSleevePlacer] MEP_ElementId parameter not found or read-only on sleeve {sleeveInstance.Id}");
                }

                sleeveInstance.LookupParameter("MEP_UniqueId")?.Set(clashZone.MepElementUniqueId);
                sleeveInstance.LookupParameter("MEP_Size")?.Set(clashZone.MepElementFormattedSize);
                sleeveInstance.LookupParameter("System_Abbreviation")?.Set(clashZone.MepElementSystemAbbreviation);
                sleeveInstance.LookupParameter("MEP_Count")?.Set(1);  // Individual sleeve

                DebugLogger.Info($"[UniversalSleevePlacer] Set parameters for sleeve {sleeveInstance.Id}");
            }
            catch (Exception ex)
            {
                DebugLogger.Warning($"[UniversalSleevePlacer] Error setting parameters: {ex.Message}");
            }
        }

        /// <summary>
        /// Set sleeve orientation based on structural element type
        /// For floors: Rotate sleeve based on MEP element direction
        /// For walls/framing: Set HostOrientation parameter based on wall normal or framing direction
        /// Mimics DuctSleevePlacer.cs lines 536-580
        /// </summary>
        private void SetSleeveOrientation(FamilyInstance sleeveInstance, ClashZone clashZone)
        {
            try
            {
                // Use pre-calculated data from ClashZone (calculated during refresh)
                // No need to get structural element and recalculate!
                bool isFloorHost = clashZone.StructuralElementType == "Floor" || 
                                  clashZone.StructuralElementType == "Floors";
                
                if (isFloorHost)
                {
                    // FLOOR: Rotate sleeve based on MEP element direction (from ClashZone)
                    var mepOrientation = clashZone.MepElementOrientation;
                    if (mepOrientation != null && (mepOrientation.X != 0 || mepOrientation.Y != 0))
                    {
                        var loc = sleeveInstance.Location as LocationPoint;
                        if (loc != null)
                        {
                            double rotationAngle = Math.Atan2(mepOrientation.Y, mepOrientation.X);
                            double rotationAngleDegrees = rotationAngle * 180 / Math.PI;
                            DebugLogger.Info($"[UniversalSleevePlacer] FLOOR: Rotating sleeve by {rotationAngleDegrees:F1}° based on MEP orientation ({mepOrientation.X:F3}, {mepOrientation.Y:F3})");
                            
                            Line rotationAxis = Line.CreateBound(loc.Point, loc.Point + XYZ.BasisZ);
                            ElementTransformUtils.RotateElement(_doc, sleeveInstance.Id, rotationAxis, rotationAngle);
                            
                            DebugLogger.Info($"[UniversalSleevePlacer] FLOOR: Applied rotation of {rotationAngleDegrees:F1}°");
                        }
                    }
                    // Also set HostOrientation parameter to "FloorHosted"
                    var hostOrientationParam = sleeveInstance.LookupParameter("HostOrientation");
                    if (hostOrientationParam != null && !hostOrientationParam.IsReadOnly)
                    {
                        hostOrientationParam.Set("FloorHosted");
                        DebugLogger.Info($"[UniversalSleevePlacer] FLOOR: Set HostOrientation = FloorHosted");
                    }
                }
                else
                {
                    // WALL or FRAMING: 
                    // 1. Rotate sleeve if MEP is Y-axis oriented
                    // 2. Set HostOrientation parameter based on wall normal
                    
                    var mepOrientation = clashZone.MepElementOrientation;
                    if (mepOrientation != null && (mepOrientation.X != 0 || mepOrientation.Y != 0))
                    {
                        // Check if MEP element is Y-axis oriented
                        bool isYAxisMep = Math.Abs(mepOrientation.Y) > Math.Abs(mepOrientation.X);
                        DebugLogger.Info($"[UniversalSleevePlacer] WALL/FRAMING: isYAxisMep={isYAxisMep}, MEP direction=({mepOrientation.X:F3},{mepOrientation.Y:F3})");
                        
                        if (isYAxisMep)
                        {
                            // Rotate sleeve 90 degrees for Y-axis MEP elements
                            var loc = sleeveInstance.Location as LocationPoint;
                            if (loc != null)
                            {
                                double rotationAngle = Math.PI / 2; // 90 degrees
                                Line rotationAxis = Line.CreateBound(loc.Point, loc.Point + XYZ.BasisZ);
                                ElementTransformUtils.RotateElement(_doc, sleeveInstance.Id, rotationAxis, rotationAngle);
                                DebugLogger.Info($"[UniversalSleevePlacer] WALL/FRAMING: Rotated Y-axis MEP sleeve 90°");
                            }
                        }
                        else
                        {
                            DebugLogger.Info($"[UniversalSleevePlacer] WALL/FRAMING: X-axis MEP - no rotation needed");
                        }
                    }
                    
                    // Set HostOrientation parameter based on wall normal
                    var structuralNormal = clashZone.StructuralElementNormal;
                    if (structuralNormal != null && (structuralNormal.X != 0 || structuralNormal.Y != 0))
                    {
                        double absX = Math.Abs(structuralNormal.X);
                        double absY = Math.Abs(structuralNormal.Y);
                        string orientation = absX > absY ? "X" : "Y";
                        
                        var hostOrientationParam = sleeveInstance.LookupParameter("HostOrientation");
                        if (hostOrientationParam != null && !hostOrientationParam.IsReadOnly)
                        {
                            hostOrientationParam.Set(orientation);
                            DebugLogger.Info($"[UniversalSleevePlacer] WALL/FRAMING: Set HostOrientation = {orientation} (wall normal: {structuralNormal.X:F3}, {structuralNormal.Y:F3})");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Warning($"[UniversalSleevePlacer] Error setting sleeve orientation: {ex.Message}");
            }
        }
    }
}
