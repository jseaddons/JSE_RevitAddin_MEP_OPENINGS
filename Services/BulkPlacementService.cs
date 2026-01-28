using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Utils;
using JSE_RevitAddin_MEP_OPENINGS.Services.Placement;

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
        private readonly SleeveParameterService _parameterService;

        public BulkPlacementService(Document doc, Action<string>? logger = null)
        {
            _logger = logger ?? (msg => DebugLogger.Info(msg));
            _parameterService = new SleeveParameterService(doc);
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
                    // ⚠️ NO NESTED TRANSACTION HERE: Orchestrator already provides one.
                    // Starting a nested transaction causes Revit to crash or elements to fail to commit.
                    var createdIds = doc.Create.NewFamilyInstances2(creationDataList);
                    var idList = createdIds.ToList();
                    _logger?.Invoke($"[BulkPlacement] Created {idList.Count} family instances");

                    // Step 4: Finalize (Rotation and Mapping)
                    for (int i = 0; i < idList.Count; i++)
                    {
                        var elementId = idList[i];
                        var zone = zoneMap[i];
                        var instance = doc.GetElement(elementId) as FamilyInstance;

                        if (instance != null)
                        {
                            ApplyRotationIfNecessary(doc, instance, zone);
                            
                            // ✅ FIX: Set Width/Height/Depth parameters from database values
                            SetClusterParameters(instance, zone);
                            
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
            // ✅ ROTATION LOGIC: Handle rotation for different host types
            // For Floors: Rotates based on MEP element direction (if circular, usually 0)
            // For Walls/Framing: Rotates to align sleeve "Width" with the host's direction
            
            try
            {
                double rotation = 0;
                XYZ axisPoint1 = (instance.Location as LocationPoint)?.Point;
                if (axisPoint1 == null) return;
                
                XYZ axisDirection = XYZ.BasisZ; // Default rotation axis for vertical hosts

                if (zone.StructuralElementType == "Floor")
                {
                    rotation = zone.MepElementRotationAngle;
                }
                else if (zone.StructuralElementType == "Wall" || zone.StructuralElementType == "Structural Framing")
                {
                    // For walls, we want to align the sleeve with the wall direction
                    // WallDirection is a normalized vector along the wall length
                    if (zone.WallDirection != null && zone.WallDirection.GetLength() > 0.001)
                    {
                        // Calculate the angle of the wall relative to the world X-axis
                        // Formula: Angle = atan2(dy, dx)
                        rotation = Math.Atan2(zone.WallDirection.Y, zone.WallDirection.X);
                        
                        // NOTE: RectangularOpeningOnWall family's default orientation 
                        // might be perpendicular to the wall or already aligned.
                        // If it's perpendicular by default, we might need a 90-degree (PI/2) offset.
                        // Based on ClusterRotationService logic, we might need to adjust this.
                    }
                }

                if (Math.Abs(rotation) > 0.001)
                {
                    XYZ axisPoint2 = axisPoint1 + axisDirection;
                    Line axis = Line.CreateBound(axisPoint1, axisPoint2);
                    instance.Location.Rotate(axis, rotation);
                }
            }
            catch (Exception ex)
            {
                _logger?.Invoke($"[BulkPlacement] Rotation failed for zone {zone.Id}: {ex.Message}");
            }
        }

        /// <summary>
        /// ✅ FIX: Set Width/Height/Depth parameters on cluster sleeve from database values
        /// Uses SleeveParameterService.SetSleeveParametersImmediate for robust transaction-safe writes.
        /// </summary>
        private void SetClusterParameters(FamilyInstance instance, ClashZone zone)
        {
            try
            {
                // Determine dimensions - use calculated values if available (for clusters)
                double width = zone.CalculatedSleeveWidth > 0 ? zone.CalculatedSleeveWidth : zone.SleeveWidth;
                double height = zone.CalculatedSleeveHeight > 0 ? zone.CalculatedSleeveHeight : zone.SleeveHeight;
                double diameter = zone.CalculatedSleeveWidth > 0 ? zone.CalculatedSleeveWidth : zone.SleeveWidth;
                
                // depth logic matches SleeveParameterService.GetThickness
                double depth = zone.CalculatedSleeveDepth > 0 ? zone.CalculatedSleeveDepth : 
                              (zone.WallThickness > 0 ? zone.WallThickness : 
                               (zone.FramingThickness > 0 ? zone.FramingThickness : zone.StructuralElementThickness));

                bool isCircular = zone.MepElementCategory == "Pipes" || zone.MepElementCategory == "Conduits";

                // ✅ CRITICAL FIX: Use immediate write path for clusters within bulk transaction
                _parameterService.SetSleeveParametersImmediate(
                    instance, 
                    width, 
                    height, 
                    diameter, 
                    isCircular, 
                    zone, 
                    depth);
                
                _parameterService.SetClusterSleeveInstanceId(instance, instance.Id.IntegerValue);
                
                _logger?.Invoke($"[BulkPlacement] Set parameters for cluster {instance.Id}: W={width*304.8:F0}mm, H={height*304.8:F0}mm, D={depth*304.8:F0}mm");
            }
            catch (Exception ex)
            {
                _logger?.Invoke($"[BulkPlacement] Failed to set parameters for zone {zone.Id}: {ex.Message}");
            }
        }
    }
}
