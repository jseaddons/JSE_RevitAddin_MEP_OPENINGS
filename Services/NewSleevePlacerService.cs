using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Services.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Services.Strategies;
using JSE_RevitAddin_MEP_OPENINGS.Utils;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// New service for placing sleeves, designed to replace UniversalSleevePlacerService.
    /// Implements SOLID principles and supports "Smart Replay" logic.
    /// </summary>
    public class NewSleevePlacerService
    {
        private readonly Document _doc;
        private readonly OpeningConditions _conditions;
        private readonly ISleevePlacementStrategy _strategy;
        private readonly Dictionary<string, double> _clearanceSettings;
        private readonly ISleeveRepository _sleeveRepository;
        private readonly IZoneFilterService _zoneFilterService;
        private readonly IFamilyManager _familyManager;
        private readonly FlagManager _flagManager; // Temporary direct dependency until extracted
        private readonly bool _isReplayPath;
        private readonly string _filterName;
        
        // Family symbol cache for performance
        private static Dictionary<string, FamilySymbol> _familySymbolCache = new Dictionary<string, FamilySymbol>();

        public NewSleevePlacerService(
            Document doc,
            OpeningConditions conditions,
            ISleevePlacementStrategy strategy,
            Dictionary<string, double> clearanceSettings,
            ISleeveRepository sleeveRepository,
            IZoneFilterService zoneFilterService,
            IFamilyManager familyManager,
            FlagManager flagManager,
            bool isReplayPath = false,
            string filterName = null)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _conditions = conditions ?? new OpeningConditions();
            _strategy = strategy ?? throw new ArgumentNullException(nameof(strategy));
            _clearanceSettings = clearanceSettings ?? new Dictionary<string, double>();
            _sleeveRepository = sleeveRepository ?? throw new ArgumentNullException(nameof(sleeveRepository));
            _zoneFilterService = zoneFilterService; // Can be null for now
            _familyManager = familyManager; // Can be null for now
            _flagManager = flagManager ?? new FlagManager(doc);
            _isReplayPath = isReplayPath;
            _filterName = filterName;
        }

        public (int placed, int skipped, int errors) PlaceAllSleevesInTransaction(List<ClashZone> clashZones)
        {
            int placed = 0;
            int skipped = 0;
            int errors = 0;
            
            var placedSleeveData = new List<(FamilyInstance sleeve, ClashZone zone)>();
            var processedZoneGuids = new List<Guid>();

            // ✅ Use injected ZoneFilterService if available to pre-filter zones
            List<ClashZone> filteredZones = clashZones;
            if (_zoneFilterService != null)
            {
                filteredZones = _zoneFilterService.PreFilterEligibleClashZones(_doc, clashZones);
                
                if (!DeploymentConfiguration.DeploymentMode && filteredZones.Count != clashZones.Count)
                {
                    DebugLogger.Info($"[NewSleevePlacer] ZoneFilterService filtered {clashZones.Count} → {filteredZones.Count} zones");
                }
            }
            
            foreach (var clashZone in filteredZones)
            {
                try
                {
                    // Skip if already resolved (unless we are forcing update, but typically we skip)
                    if (clashZone.IsResolved || clashZone.IsClusterResolved)
                    {
                        skipped++;
                        continue;
                    }

                    bool isSleevePlaced = false;
                    FamilyInstance placedSleeve = null;

                    // SMART REPLAY LOGIC (Path 1)
                    if (_isReplayPath)
                    {
                        // Check if sleeve exists in Revit
                        bool sleeveExists = false;
                        if (clashZone.SleeveInstanceId > 0)
                        {
                            var element = _doc.GetElement(new ElementId(clashZone.SleeveInstanceId));
                            if (element != null && element is FamilyInstance)
                            {
                                sleeveExists = true;
                            }
                        }

                        if (sleeveExists)
                        {
                            // Sleeve exists and we are in replay path -> Assume it's correct and skip (or update if needed)
                            // For this refactor, we'll skip placement but mark as processed
                            skipped++;
                            processedZoneGuids.Add(clashZone.Id);
                            continue;
                        }
                        else
                        {
                            // Sleeve is missing (deleted) but we have saved data
                            // Check if we can use saved data ("Smart Replay")
                            if (CanUseSavedData(clashZone))
                            {
                                placedSleeve = PlaceSleeveFromSavedData(clashZone);
                                if (placedSleeve != null)
                                {
                                    isSleevePlaced = true;
                                    placed++;
                                }
                            }
                        }
                    }

                    // If not placed via Smart Replay, proceed with Normal Placement
                    if (!isSleevePlaced)
                    {
                        placedSleeve = PlaceSleeveNormal(clashZone);
                        if (placedSleeve != null)
                        {
                            placed++;
                        }
                        else
                        {
                            skipped++; // Failed to place for some reason (e.g. invalid dimensions)
                        }
                    }

                    if (placedSleeve != null)
                    {
                        // Update ClashZone with new Sleeve ID
                        clashZone.SleeveInstanceId = placedSleeve.Id.IntegerValue;
                        clashZone.IsResolved = true;
                        
                        placedSleeveData.Add((placedSleeve, clashZone));
                        processedZoneGuids.Add(clashZone.Id);
                        
                        // Update Database immediately (or batch later - doing per-sleeve for safety now)
                        if (OptimizationFlags.UseNewSleeveRepository)
                        {
                            UpdateSleeveDataInDatabase(clashZone, placedSleeve);
                        }
                    }
                }
                catch (Exception ex)
                {
                    errors++;
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Error($"[NewSleevePlacer] Error placing sleeve for zone {clashZone.Id}: {ex.Message}");
                }
            }

            // Batch update flags at the end
            if (placedSleeveData.Count > 0)
            {
                try
                {
                    var batchUpdates = placedSleeveData
                        .Where(x => x.zone.SleeveInstanceId > 0)
                        .Select(x => (x.zone, x.zone.SleeveInstanceId))
                        .ToList();

                    if (batchUpdates.Count > 0)
                    {
                        _flagManager.BatchUpdateFlagsForPlacement(
                            batchUpdates,
                            isCluster: false,
                            placedSleeveData[0].zone.MepElementCategory,
                            _filterName
                        );
                    }
                }
                catch (Exception ex)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Error($"[NewSleevePlacer] Error updating flags: {ex.Message}");
                }
            }
            
            // Reset ReadyForPlacement flags
            if (processedZoneGuids.Count > 0)
            {
                ResetReadyForPlacementFlags(processedZoneGuids);
            }

            return (placed, skipped, errors);
        }

        private bool CanUseSavedData(ClashZone zone)
        {
            // Check if we have valid saved dimensions and placement point
            // Also verify clearance hasn't changed (simplified check for now)
            return zone.SleeveWidth > 0 && 
                   zone.SleeveHeight > 0 && 
                   zone.SleevePlacementPointX != 0;
        }

        private FamilyInstance PlaceSleeveFromSavedData(ClashZone zone)
        {
            // Use saved dimensions directly
            double width = zone.SleeveWidth;
            double height = zone.SleeveHeight;
            double diameter = zone.SleeveDiameter;
            
            // Determine shape based on dimensions
            bool isCircular = diameter > 0 && width <= 0;
            
            // Select Family
            string familyName = GetSleeveFamilyName(zone, isCircular);
            FamilySymbol symbol = LoadFamilySymbol(familyName);
            
            if (symbol == null) return null;

            // Place Instance
            XYZ placementPoint = new XYZ(zone.SleevePlacementPointX, zone.SleevePlacementPointY, zone.SleevePlacementPointZ);
            
            // Use saved rotation if available, otherwise default to MEP rotation
            double rotation = zone.MepElementRotationAngle; 

            FamilyInstance instance = PlaceSleeveInstance(symbol, placementPoint, zone, rotation);
            
            if (instance != null)
            {
                SetSleeveParameters(instance, width, height, diameter, isCircular);
            }
            
            return instance;
        }

        private FamilyInstance PlaceSleeveNormal(ClashZone zone)
        {
            // Calculate dimensions based on MEP element and clearance using strategy
            var (width, height, diameter, isCircular) = CalculateSleeveDimensions(zone);
            
            if (width <= 0 && height <= 0 && diameter <= 0) return null; // Invalid dimensions

            // Select Family
            string familyName = GetSleeveFamilyName(zone, isCircular);
            FamilySymbol symbol = LoadFamilySymbol(familyName);
            
            if (symbol == null) return null;

            // Determine Placement Point (Intersection Point)
            XYZ placementPoint = zone.IntersectionPoint ?? new XYZ(zone.IntersectionPointX, zone.IntersectionPointY, zone.IntersectionPointZ);
            
            // Determine Rotation
            double rotation = zone.MepElementRotationAngle;

            // Place Instance
            FamilyInstance instance = PlaceSleeveInstance(symbol, placementPoint, zone, rotation);
            
            if (instance != null)
            {
                SetSleeveParameters(instance, width, height, diameter, isCircular);
                
                // Update zone with calculated dimensions for saving
                zone.SleeveWidth = width;
                zone.SleeveHeight = height;
                zone.SleeveDiameter = diameter;
                zone.SleevePlacementPoint = placementPoint;
                zone.SleevePlacementPointX = placementPoint.X;
                zone.SleevePlacementPointY = placementPoint.Y;
                zone.SleevePlacementPointZ = placementPoint.Z;
            }
            
            return instance;
        }

        private (double width, double height, double diameter, bool isCircular) CalculateSleeveDimensions(ClashZone zone)
        {
            // Use strategy to calculate clearance
            double clearance = GetClearance(zone);
            
            double width = 0;
            double height = 0;
            double diameter = 0;
            bool isCircular = false;

            if (zone.MepElementDiameter > 0)
            {
                // Circular MEP
                isCircular = true;
                diameter = zone.MepElementDiameter + (2 * clearance);
            }
            else
            {
                // Rectangular MEP
                isCircular = false;
                width = zone.MepElementWidth + (2 * clearance);
                height = zone.MepElementHeight + (2 * clearance);
            }

            return (width, height, diameter, isCircular);
        }

        private double GetClearance(ClashZone zone)
        {
            // Use strategy to calculate clearance based on conditions
            try
            {
                // Get MEP element size for strategy
                double mepSize = zone.MepElementDiameter > 0 
                    ? zone.MepElementDiameter 
                    : Math.Max(zone.MepElementWidth, zone.MepElementHeight);

                // Use strategy to calculate clearance
                var clearanceResult = _strategy.CalculateClearance(
                    zone.StructuralElementType,
                    zone.MepElementCategory,
                    mepSize,
                    _conditions
                );

                return clearanceResult.clearance;
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[NewSleevePlacer] Strategy clearance calculation failed for zone {zone.Id}, using fallback: {ex.Message}");
                
                // Fallback to clearance settings or default
                string key = $"{zone.StructuralElementType}_{zone.MepElementCategory}";
                if (_clearanceSettings.ContainsKey(key))
                {
                    return _clearanceSettings[key];
                }
                
                // Final fallback: 50mm default
                return RevitUnitConversionService.Instance.ToInternalMillimeters(50);
            }
        }

        private string GetSleeveFamilyName(ClashZone zone, bool isCircular)
        {
            // Determine family based on host type (Wall vs Slab) and shape
            bool isWall = zone.StructuralElementType.Contains("Wall", StringComparison.OrdinalIgnoreCase);
            
            if (isWall)
            {
                return isCircular ? "OpeningOnWall-Circular" : "OpeningOnWall-Rectangular";
            }
            else // Slab/Floor
            {
                return isCircular ? "OpeningOnSlab-Circular" : "OpeningOnSlab-Rectangular";
            }
        }

        private FamilySymbol LoadFamilySymbol(string familyName)
        {
            // Check cache first
            if (_familySymbolCache.ContainsKey(familyName))
            {
                return _familySymbolCache[familyName];
            }
            
            // Use FamilyManager if available
            if (_familyManager != null)
            {
                var symbol = _familyManager.LoadFamily(familyName);
                if (symbol != null)
                {
                    _familySymbolCache[familyName] = symbol;
                    return symbol;
                }
            }
            
            // Fallback to simple lookup
            var foundSymbol = new FilteredElementCollector(_doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault(x => x.FamilyName.Equals(familyName, StringComparison.OrdinalIgnoreCase));
            
            if (foundSymbol != null)
            {
                _familySymbolCache[familyName] = foundSymbol;
            }
            
            return foundSymbol;
        }

        private FamilyInstance PlaceSleeveInstance(FamilySymbol symbol, XYZ point, ClashZone zone, double rotation)
        {
            if (!symbol.IsActive) symbol.Activate();

            // Determine level
            Level level = null;
            if (zone.LevelId > 0)
            {
                level = _doc.GetElement(new ElementId(zone.LevelId)) as Level;
            }
            
            if (level == null)
            {
                // Fallback to nearest level or active view level
                level = _doc.ActiveView.GenLevel;
            }

            // Place instance
            FamilyInstance instance = _doc.Create.NewFamilyInstance(point, symbol, level, StructuralType.NonStructural);
            
            // Apply rotation
            if (Math.Abs(rotation) > 1e-6)
            {
                Line axis = Line.CreateBound(point, point + XYZ.BasisZ);
                ElementTransformUtils.RotateElement(_doc, instance.Id, axis, rotation);
            }

            return instance;
        }

        private void SetSleeveParameters(FamilyInstance instance, double width, double height, double diameter, bool isCircular)
        {
            if (isCircular)
            {
                var param = instance.LookupParameter("Diameter") ?? instance.LookupParameter("Sleeve Diameter");
                if (param != null && !param.IsReadOnly) param.Set(diameter);
            }
            else
            {
                var wParam = instance.LookupParameter("Width") ?? instance.LookupParameter("Sleeve Width");
                var hParam = instance.LookupParameter("Height") ?? instance.LookupParameter("Sleeve Height");
                if (wParam != null && !wParam.IsReadOnly) wParam.Set(width);
                if (hParam != null && !hParam.IsReadOnly) hParam.Set(height);
            }
        }

        private void UpdateSleeveDataInDatabase(ClashZone zone, FamilyInstance sleeve)
        {
            try
            {
                using (var context = new SleeveDbContext(_doc, msg => { }))
                {
                    var repository = new ClashZoneRepository(context, msg => { });
                    
                    // Update Instance ID
                    repository.UpdateSleeveInstanceId(zone.Id, sleeve.Id.IntegerValue);
                    
                    // Update Placement Data
                    repository.UpdateSleevePlacement(
                        zone.Id,
                        sleeve.Id.IntegerValue,
                        zone.SleeveWidth,
                        zone.SleeveHeight,
                        zone.SleeveDiameter,
                        zone.SleevePlacementPointX,
                        zone.SleevePlacementPointY,
                        zone.SleevePlacementPointZ,
                        zone.SleevePlacementPointX, // Active doc coords (same for now)
                        zone.SleevePlacementPointY,
                        zone.SleevePlacementPointZ,
                        zone.MepElementRotationAngle
                    );
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[NewSleevePlacer] Failed to update DB for zone {zone.Id}: {ex.Message}");
            }
        }
        
        private void ResetReadyForPlacementFlags(List<Guid> zoneIds)
        {
            try
            {
                using (var context = new SleeveDbContext(_doc, msg => { }))
                {
                    var repository = new ClashZoneRepository(context, msg => { });
                    repository.BulkResetReadyForPlacementFlags(zoneIds);
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[NewSleevePlacer] Failed to reset flags: {ex.Message}");
            }
        }
    }
}
