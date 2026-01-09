using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Result class for bulk placement operations.
    /// Tracks success/failure counts and provides user-friendly summaries.
    /// </summary>
    public class BulkPlacementResult
    {
        public bool OverallSuccess { get; set; }
        public int TotalZones { get; set; }
        public int PlacedCount { get; set; }
        public int SkippedCount { get; set; }
        public int FailedCount { get; set; }
        public string? Error { get; set; }
        public string? Warning { get; set; }
        public List<(Guid ZoneGuid, string Reason)> Failures { get; set; } = new();
        public TimeSpan ElapsedTime { get; set; }
        
        public string GetSummary()
        {
            if (OverallSuccess)
                return $"✅ Successfully placed {PlacedCount}/{TotalZones} sleeves in {ElapsedTime.TotalSeconds:F1}s";
            else if (PlacedCount > 0)
                return $"⚠️ Partial success: {PlacedCount} placed, {FailedCount} failed. {Error}";
            else
                return $"❌ Failed: {Error}";
        }
    }

    /// <summary>
    /// Interface for bulk placement service (SOLID DIP).
    /// </summary>
    public interface IBulkPlacementService
    {
        BulkPlacementResult ExecuteBulkPlacement(Document doc);
    }

    /// <summary>
    /// ✅ BULK PLACEMENT SERVICE: Places individual sleeves using NewFamilyInstances2.
    /// Uses OptimizationFlags.UseBulkIndividualSleevePlacement flag.
    /// Scope: Individual sleeves only (cluster sleeves in Phase 2).
    /// </summary>
    public class BulkPlacementService : IBulkPlacementService
    {
        private readonly SleeveDbContext _dbContext;
        private readonly Action<string>? _logger;
        
        public BulkPlacementService(SleeveDbContext dbContext, Action<string>? logger = null)
        {
            _dbContext = dbContext;
            _logger = logger ?? (msg => DebugLogger.Info(msg));
        }

        /// <summary>
        /// Execute bulk placement for all zones ready for placement.
        /// 7-step flow as per Bulk_Family_Placement_Strategy.md.
        /// </summary>
        public BulkPlacementResult ExecuteBulkPlacement(Document doc)
        {
            var startTime = DateTime.Now;
            var result = new BulkPlacementResult();
            
            try
            {
                // Step 1: Query DB
                var repository = new ClashZoneRepository(_dbContext);
                var zones = repository.GetZonesReadyForPlacement();
                result.TotalZones = zones.Count;
                
                if (zones.Count == 0)
                {
                    result.OverallSuccess = true;
                    result.ElapsedTime = DateTime.Now - startTime;
                    _logger?.Invoke("[BulkPlacement] No zones ready for placement");
                    return result;
                }
                
                _logger?.Invoke($"[BulkPlacement] Step 1: Retrieved {zones.Count} zones for bulk placement");
                
                // Step 2: Pre-flight validation
                ValidateBeforePlacement(zones);
                _logger?.Invoke($"[BulkPlacement] Step 2: Pre-flight validation passed");
                
                // Step 3: Calculate sizes and family names (using current UI settings)
                CalculateSizesAndFamilyNames(zones, doc);
                _logger?.Invoke($"[BulkPlacement] Step 3: Calculated sizes for {zones.Count} zones");
                
                // Step 4: Save calculated sizes to DB
                repository.BatchUpdateSleeveSizesForBulkPlacement(zones);
                _logger?.Invoke($"[BulkPlacement] Step 4: Saved sizes to database");
                
                // Steps 5-7: All Revit changes in TransactionGroup
                using (var txGroup = new TransactionGroup(doc, "Bulk Sleeve Placement"))
                {
                    txGroup.Start();
                    
                    try
                    {
                        // Step 5: Pre-activate symbols
                        Dictionary<string, FamilySymbol> symbolCache;
                        using (var tx = new Transaction(doc, "Activate Symbols"))
                        {
                            tx.Start();
                            symbolCache = PreActivateSymbols(doc, zones);
                            tx.Commit();
                        }
                        _logger?.Invoke($"[BulkPlacement] Step 5: Activated {symbolCache.Count} unique symbols");
                        
                        // Step 6: Build creation data and place
                        using (var tx = new Transaction(doc, "Place Sleeves"))
                        {
                            tx.Start();
                            
                            // Build FamilyInstanceCreationData list
                            var creationDataList = new List<Autodesk.Revit.Creation.FamilyInstanceCreationData>();
                            foreach (var zone in zones)
                            {
                                if (symbolCache.TryGetValue(zone.SleeveFamilyName, out var symbol))
                                {
                                    var point = new XYZ(
                                        zone.SleevePlacementPointX,
                                        zone.SleevePlacementPointY,
                                        zone.SleevePlacementPointZ);
                                    
                                    var creationData = new Autodesk.Revit.Creation.FamilyInstanceCreationData(
                                        point, symbol, StructuralType.NonStructural);
                                    creationDataList.Add(creationData);
                                }
                                else
                                {
                                    result.Failures.Add((zone.Id, $"Symbol not found: {zone.SleeveFamilyName}"));
                                }
                            }
                            
                            // Bulk placement
                            var createdIds = doc.Create.NewFamilyInstances2(creationDataList);
                            var idList = createdIds.ToList();
                            
                            // Set parameters on created instances
                            int successIndex = 0;
                            for (int i = 0; i < zones.Count; i++)
                            {
                                var zone = zones[i];
                                if (!symbolCache.ContainsKey(zone.SleeveFamilyName))
                                {
                                    result.FailedCount++;
                                    continue;
                                }
                                
                                var sleeve = doc.GetElement(idList[successIndex]) as FamilyInstance;
                                if (sleeve != null)
                                {
                                    // Set parameters
                                    SetSleeveParameters(sleeve, zone, result);
                                    
                                    // Apply rotation for floors
                                    ApplyRotation(doc, sleeve, zone, result);
                                    
                                    // Store instance ID for DB update
                                    zone.SleeveInstanceId = sleeve.Id.IntegerValue;
                                    result.PlacedCount++;
                                }
                                else
                                {
                                    result.Failures.Add((zone.Id, "Created element was null"));
                                    result.FailedCount++;
                                }
                                successIndex++;
                            }
                            
                            doc.Regenerate();
                            tx.Commit();
                        }
                        
                        _logger?.Invoke($"[BulkPlacement] Step 6: Placed {result.PlacedCount} sleeves");
                        
                        txGroup.Assimilate();
                    }
                    catch (Exception ex)
                    {
                        txGroup.RollBack();
                        result.Error = $"Placement failed: {ex.Message}";
                        result.OverallSuccess = false;
                        result.ElapsedTime = DateTime.Now - startTime;
                        _logger?.Invoke($"[BulkPlacement] ❌ Error during placement: {ex.Message}");
                        return result;
                    }
                }
                
                // Step 7: Update DB with retry logic
                BatchUpdateDatabaseWithRetry(repository, zones, result);
                _logger?.Invoke($"[BulkPlacement] Step 7: Updated database with {result.PlacedCount} sleeve IDs");
                
                result.OverallSuccess = result.FailedCount == 0;
                result.ElapsedTime = DateTime.Now - startTime;
                _logger?.Invoke($"[BulkPlacement] ✅ Complete: {result.GetSummary()}");
            }
            catch (Exception ex)
            {
                result.Error = ex.Message;
                result.OverallSuccess = false;
                result.ElapsedTime = DateTime.Now - startTime;
                _logger?.Invoke($"[BulkPlacement] ❌ Fatal error: {ex.Message}");
            }
            
            return result;
        }

        private void ValidateBeforePlacement(List<ClashZone> zones)
        {
            var errors = new List<string>();
            
            // Check for invalid coordinates
            var invalidCoords = zones.Count(z => 
                double.IsNaN(z.SleevePlacementPointX) ||
                double.IsNaN(z.SleevePlacementPointY) ||
                double.IsNaN(z.SleevePlacementPointZ) ||
                double.IsInfinity(z.SleevePlacementPointX));
            if (invalidCoords > 0)
                errors.Add($"Invalid coordinates in {invalidCoords} zones");
            
            // Check for duplicate zones
            var duplicates = zones.GroupBy(z => z.Id).Where(g => g.Count() > 1).Count();
            if (duplicates > 0)
                errors.Add($"Duplicate zones detected: {duplicates}");
            
            if (errors.Any())
                throw new InvalidOperationException($"Pre-flight validation failed:\n• {string.Join("\n• ", errors)}");
        }

        private void CalculateSizesAndFamilyNames(List<ClashZone> zones, Document doc)
        {
            // TODO: Implement using existing CalculateSleeveDimensions and DetermineOpeningType logic
            // For now, use pre-populated values from database
            // This will be wired up to NewSleevePlacerService logic in Phase 3 integration
            
            foreach (var zone in zones)
            {
                // Ensure family name is set (should already be populated from Refresh)
                if (string.IsNullOrEmpty(zone.SleeveFamilyName))
                {
                    zone.SleeveFamilyName = GetDefaultFamilyName(zone);
                }
            }
        }

        private string GetDefaultFamilyName(ClashZone zone)
        {
            bool isWallOrFraming = zone.StructuralElementType == "Wall" || 
                                  zone.StructuralElementType == "Walls" ||
                                  zone.StructuralElementType == "Structural Framing";
            
            // Default to rectangular for bulk placement
            return isWallOrFraming ? "RectangularOpeningOnWall" : "RectangularOpeningOnSlab";
        }

        private Dictionary<string, FamilySymbol> PreActivateSymbols(Document doc, List<ClashZone> zones)
        {
            var uniqueFamilyNames = zones.Select(z => z.SleeveFamilyName)
                .Where(n => !string.IsNullOrEmpty(n))
                .Distinct().ToList();
            
            var cache = new Dictionary<string, FamilySymbol>();
            var missingFamilies = new List<string>();
            
            // Phase 1: Validate all exist
            foreach (var familyName in uniqueFamilyNames)
            {
                var family = new FilteredElementCollector(doc)
                    .OfClass(typeof(Family))
                    .Cast<Family>()
                    .FirstOrDefault(f => f.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase));
                
                if (family == null)
                {
                    missingFamilies.Add(familyName);
                    continue;
                }
                
                var symbolIds = family.GetFamilySymbolIds();
                if (symbolIds == null || symbolIds.Count == 0)
                    missingFamilies.Add(familyName);
            }
            
            if (missingFamilies.Count > 0)
            {
                throw new InvalidOperationException(
                    $"Cannot proceed: Missing families [{string.Join(", ", missingFamilies)}]. " +
                    "Load these families before placement.");
            }
            
            // Phase 2: Activate
            foreach (var familyName in uniqueFamilyNames)
            {
                var family = new FilteredElementCollector(doc)
                    .OfClass(typeof(Family))
                    .Cast<Family>()
                    .First(f => f.Name.Equals(familyName, StringComparison.OrdinalIgnoreCase));
                
                var symbolId = family.GetFamilySymbolIds().First();
                var symbol = doc.GetElement(symbolId) as FamilySymbol;
                
                if (!symbol.IsActive)
                    symbol.Activate();
                
                cache[familyName] = symbol;
            }
            
            return cache;
        }

        private void SetSleeveParameters(FamilyInstance sleeve, ClashZone zone, BulkPlacementResult result)
        {
            try
            {
                sleeve.LookupParameter("Width")?.Set(zone.SleeveWidth);
                sleeve.LookupParameter("Height")?.Set(zone.SleeveHeight);
                if (zone.SleeveDiameter > 0)
                    sleeve.LookupParameter("Diameter")?.Set(zone.SleeveDiameter);
            }
            catch (Exception ex)
            {
                result.Failures.Add((zone.Id, $"Parameter setting failed: {ex.Message}"));
            }
        }

        private void ApplyRotation(Document doc, FamilyInstance sleeve, ClashZone zone, BulkPlacementResult result)
        {
            if (!string.Equals(zone.StructuralElementType, "Floor", StringComparison.OrdinalIgnoreCase))
                return;
            
            if (Math.Abs(zone.MepElementRotationAngle) < 0.001)
                return;
            
            try
            {
                var locationPoint = sleeve.Location as LocationPoint;
                if (locationPoint == null) return;
                
                var point = locationPoint.Point;
                var axis = Line.CreateBound(point, point.Add(XYZ.BasisZ));
                ElementTransformUtils.RotateElement(doc, sleeve.Id, axis, zone.MepElementRotationAngle);
            }
            catch (Exception ex)
            {
                result.Failures.Add((zone.Id, $"Rotation failed: {ex.Message}"));
            }
        }

        private void BatchUpdateDatabaseWithRetry(ClashZoneRepository repository, List<ClashZone> zones, BulkPlacementResult result)
        {
            int retryCount = 0;
            const int maxRetries = 3;
            
            while (retryCount < maxRetries)
            {
                try
                {
                    repository.BatchUpdateAfterBulkPlacement(zones);
                    return;
                }
                catch (Exception ex)
                {
                    retryCount++;
                    _logger?.Invoke($"[BulkPlacement] ⚠️ DB update attempt {retryCount} failed: {ex.Message}");
                    
                    if (retryCount >= maxRetries)
                    {
                        // Save recovery file
                        var recoveryData = zones.Select(z => new {
                            ZoneGuid = z.Id,
                            SleeveInstanceId = z.SleeveInstanceId
                        });
                        
                        string recoveryPath = Path.Combine(
                            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                            "JSE_MEP_Openings", "Logs",
                            $"placement_recovery_{DateTime.Now:yyyyMMdd_HHmmss}.json");
                        
                        try
                        {
                            Directory.CreateDirectory(Path.GetDirectoryName(recoveryPath)!);
                            File.WriteAllText(recoveryPath, JsonSerializer.Serialize(recoveryData, new JsonSerializerOptions { WriteIndented = true }));
                            _logger?.Invoke($"[BulkPlacement] Recovery file saved to: {recoveryPath}");
                        }
                        catch { /* Ignore file write errors */ }
                        
                        result.Warning = $"Sleeves placed but DB update failed. Recovery file: {recoveryPath}";
                        return;
                    }
                    
                    System.Threading.Thread.Sleep(500 * retryCount);
                }
            }
        }
    }
}
