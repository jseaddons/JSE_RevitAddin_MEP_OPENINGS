using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Utils;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    public class BulkPlacementResult
    {
        public bool OverallSuccess { get; set; }
        public int TotalZones { get; set; }
        public int PlacedCount { get; set; }
        public int FailedCount { get; set; }
        public string Error { get; set; }
        public List<(ClashZone Zone, ElementId ElementId)> PlacedItems { get; set; } = new();
        public List<(Guid ZoneGuid, string Reason)> Failures { get; set; } = new();
        public TimeSpan ElapsedTime { get; set; }
    }

    public interface IBulkPlacementService
    {
        BulkPlacementResult ExecuteBulkPlacement(Document doc, List<ClashZone> zones);
    }

    /// <summary>
    /// PURE REVIT API SERVICE: Handles batch placement of families.
    /// Does NOT interact with the database or XML files.
    /// </summary>
    public class BulkPlacementService : IBulkPlacementService
    {
        private readonly Action<string>? _logger;

        public BulkPlacementService(Action<string>? logger = null)
        {
            _logger = logger ?? (msg => DebugLogger.Info(msg));
        }

        public BulkPlacementResult ExecuteBulkPlacement(Document doc, List<ClashZone> zones)
        {
            var startTime = DateTime.Now;
            var result = new BulkPlacementResult { TotalZones = zones.Count };

            if (zones == null || zones.Count == 0)
            {
                result.OverallSuccess = true;
                return result;
            }

            try
            {
                // Step 1: Pre-activate Symbols
                var symbolCache = PreActivateSymbols(doc, zones);
                _logger?.Invoke($"[BulkPlacement] Activated {symbolCache.Count} unique symbols");

                // Step 2: Build Creation Data
                var creationDataList = new List<Autodesk.Revit.Creation.FamilyInstanceCreationData>();
                var zoneMap = new List<ClashZone>();

                foreach (var zone in zones)
                {
                    if (string.IsNullOrEmpty(zone.SleeveFamilyName)) continue;

                    if (symbolCache.TryGetValue(zone.SleeveFamilyName, out var symbol))
                    {
                        // Use IntersectionPoint for correct wall sleeve placement
                        // SleevePlacementPoint is pre-populated with WallCenterline which causes "half in, half out"
                        var point = new XYZ(
                            zone.IntersectionPointX,
                            zone.IntersectionPointY,
                            zone.IntersectionPointZ);

                        var creationData = new Autodesk.Revit.Creation.FamilyInstanceCreationData(
                            point, symbol, StructuralType.NonStructural);
                        
                        creationDataList.Add(creationData);
                        zoneMap.Add(zone);
                    }
                    else
                    {
                        result.Failures.Add((zone.Id, $"Symbol not found: {zone.SleeveFamilyName}"));
                        result.FailedCount++;
                    }
                }

                if (creationDataList.Count > 0)
                {
                    // Step 3: Bulk Placement
                    var createdIds = doc.Create.NewFamilyInstances2(creationDataList);
                    var idList = createdIds.ToList();

                    // Step 4: Finalize (Rotation and Mapping)
                    for (int i = 0; i < idList.Count; i++)
                    {
                        var elementId = idList[i];
                        var zone = zoneMap[i];
                        var instance = doc.GetElement(elementId) as FamilyInstance;

                        if (instance != null)
                        {
                            ApplyRotationIfNecessary(doc, instance, zone);
                            result.PlacedItems.Add((zone, elementId));
                            result.PlacedCount++;
                        }
                        else
                        {
                            result.Failures.Add((zone.Id, "Created instance was null"));
                            result.FailedCount++;
                        }
                    }
                }

                result.OverallSuccess = true;
                result.ElapsedTime = DateTime.Now - startTime;
                _logger?.Invoke($"[BulkPlacement] Successfully placed {result.PlacedCount} sleeves in {result.ElapsedTime.TotalSeconds:F2}s");
            }
            catch (Exception ex)
            {
                result.OverallSuccess = false;
                result.Error = ex.Message;
                _logger?.Invoke($"[BulkPlacement] CRITICAL ERROR: {ex.Message}");
            }

            return result;
        }

        private Dictionary<string, FamilySymbol> PreActivateSymbols(Document doc, List<ClashZone> zones)
        {
            var cache = new Dictionary<string, FamilySymbol>();
            var familyNames = zones.Select(z => z.SleeveFamilyName).Distinct().Where(f => !string.IsNullOrEmpty(f));

            foreach (var familyName in familyNames)
            {
                var symbol = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>()
                    .FirstOrDefault(s => s.Family.Name == familyName);

                if (symbol != null)
                {
                    if (!symbol.IsActive) symbol.Activate();
                    cache[familyName] = symbol;
                }
            }

            return cache;
        }

        private void ApplyRotationIfNecessary(Document doc, FamilyInstance instance, ClashZone zone)
        {
            // Floor sleeves often need rotation based on the MEP element's direction
            if (zone.StructuralElementType == "Floor")
            {
                try
                {
                    double rotation = zone.MepElementRotationAngle;
                    if (Math.Abs(rotation) > 0.001)
                    {
                        XYZ axisPoint1 = (instance.Location as LocationPoint).Point;
                        XYZ axisPoint2 = axisPoint1 + XYZ.BasisZ;
                        Line axis = Line.CreateBound(axisPoint1, axisPoint2);
                        instance.Location.Rotate(axis, rotation);
                    }
                }
                catch (Exception ex)
                {
                    _logger?.Invoke($"[BulkPlacement] Rotation failed for zone {zone.Id}: {ex.Message}");
                }
            }
        }
    }
}
