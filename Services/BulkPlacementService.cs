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
        BulkPlacementResult ExecuteBulkPlacement(Document doc, List<(ClashZone Zone, SleevePlacementPlanningDto Plan)> taskItems);
    }

    /// <summary>
    /// PURE REVIT API SERVICE: Handles batch placement of families.
    /// Does NOT interact with the database or XML files.
    /// 
    /// ✅ SIMPLIFIED: Uses plan.PlacementPoint directly from ParallelSleevePlacementPlanner.
    /// The planner calculates all placement points including damper offsets and rounding.
    /// </summary>
    public class BulkPlacementService : IBulkPlacementService
    {
        private readonly Action<string>? _logger;
        private readonly SleeveParameterService _parameterService;
        private readonly SleeveRotationService _rotationService;
        private readonly JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.IPerformanceMonitor? _performanceMonitor;

        public BulkPlacementService(
            Document doc,
            Action<string>? logger = null,
            JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.IPerformanceMonitor? performanceMonitor = null,
            SleeveParameterService? parameterService = null,
            SleeveRotationService? rotationService = null)
        {
            _logger = logger ?? (msg => DebugLogger.Info(msg));
            _parameterService = parameterService ?? new SleeveParameterService(doc);
            _rotationService = rotationService ?? new SleeveRotationService();
            _performanceMonitor = performanceMonitor;
        }

        public BulkPlacementResult ExecuteBulkPlacement(Document doc, List<(ClashZone Zone, SleevePlacementPlanningDto Plan)> taskItems)
        {
            var startTime = DateTime.Now;
            var result = new BulkPlacementResult { TotalZones = taskItems.Count };

            if (taskItems == null || taskItems.Count == 0)
            {
                result.OverallSuccess = true;
                return result;
            }

            try
            {
                // Step 1: Pre-activate Symbols
                Dictionary<string, FamilySymbol> symbolCache;
                using (_performanceMonitor?.TrackOperation("Pre-activate Symbols"))
                {
                    symbolCache = PreActivateSymbols(doc, taskItems.Select(x => x.Plan).ToList());
                    _logger?.Invoke($"[BulkPlacement] Activated {symbolCache.Count} unique symbols");
                }

                // Step 2: Build Creation Data (with safety dedupe by location)
                var creationDataList = new List<Autodesk.Revit.Creation.FamilyInstanceCreationData>();
                var itemMap = new List<(ClashZone Zone, SleevePlacementPlanningDto Plan)>();

                const double locationToleranceFt = 0.00656; // ~2mm
                double RoundLoc(double v) => Math.Round(v / locationToleranceFt) * locationToleranceFt;
                var seenLocations = new HashSet<(string fam, double x, double y, double z)>();

                using (_performanceMonitor?.TrackOperation("Build Creation Data"))
                {
                    foreach (var item in taskItems)
                    {
                        var zone = item.Zone;
                        var plan = item.Plan;

                        if (string.IsNullOrEmpty(plan.SleeveFamilyName)) continue;

                        if (symbolCache.TryGetValue(plan.SleeveFamilyName, out var symbol))
                        {
                            // ✅ SIMPLIFIED: Use plan.PlacementPoint directly!
                            // ParallelSleevePlacementPlanner calculates the correct point including:
                            // - Damper connector-side offset: (MEP Clearance - Other Clearance) / 2 = 25mm for 100/50
                            // - Wall centerline adjustment
                            // - All other adjustments
                            XYZ point = plan.PlacementPoint;
                            
                            // Fallback only if PlacementPoint is null (should not happen with proper planning)
                            if (point == null)
                            {
                                point = new XYZ(zone.IntersectionPointX, zone.IntersectionPointY, zone.IntersectionPointZ);
                                _logger?.Invoke($"[BulkPlacement] ⚠️ WARNING: plan.PlacementPoint was null for zone {zone.Id}, using intersection point");
                            }

                            var key = (plan.SleeveFamilyName ?? zone.SleeveFamilyName ?? "", RoundLoc(point.X), RoundLoc(point.Y), RoundLoc(point.Z));
                            if (seenLocations.Contains(key))
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                    _logger?.Invoke($"[BulkPlacement] Skipped duplicate location for zone {zone.Id} at ({point.X:F4}, {point.Y:F4}, {point.Z:F4})");
                                continue;
                            }
                            seenLocations.Add(key);

                            var creationData = new Autodesk.Revit.Creation.FamilyInstanceCreationData(
                                point, symbol, StructuralType.NonStructural);

                            creationDataList.Add(creationData);
                            itemMap.Add(item);
                        }
                        else
                        {
                            result.Failures.Add((zone.Id, $"Symbol not found: {plan.SleeveFamilyName}"));
                            result.FailedCount++;
                        }
                    }
                }

                if (creationDataList.Count > 0)
                {
                    // Step 3: Bulk Placement
                    ICollection<ElementId> createdIds;
                    using (_performanceMonitor?.TrackOperation("Revit NewFamilyInstances2"))
                    {
                        createdIds = doc.Create.NewFamilyInstances2(creationDataList);
                        
                        var idListDiag = createdIds.ToList();
                        SafeFileLogger.SafeAppendText("placement_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] [BULK-PLACEMENT-RESULT] createdIds.Count = {idListDiag.Count}\n");
                    }
                    var idList = createdIds.ToList();
                    _logger?.Invoke($"[BulkPlacement] Created {idList.Count} family instances");

                    // Step 4: Finalize (Rotation and Parameters)
                    using (_performanceMonitor?.TrackOperation("Apply Rotation & Parameters"))
                    {
                        for (int i = 0; i < idList.Count; i++)
                        {
                            var elementId = idList[i];
                            var item = itemMap[i];
                            var instance = doc.GetElement(elementId) as FamilyInstance;

                            if (instance != null)
                            {
                                bool isWallOrFraming = IsWallOrFraming(item.Zone);

                                if (isWallOrFraming)
                                {
                                    // WALL/FRAMING PATH: Rotation from RotationService
                                    double rotationRad = _rotationService.DetermineRotation(item.Zone);
                                    ApplyRotation(doc, instance, rotationRad);
                                    SetWallFramingParameters(instance, item.Zone, item.Plan);
                                }
                                else
                                {
                                    // FLOOR PATH: Rotation from Plan directly
                                    ApplyRotation(doc, instance, item.Plan.RotationRadians);
                                    SetFloorParameters(instance, item.Zone, item.Plan);
                                }

                                _parameterService.SetSleeveInstanceId(instance, elementId.IntegerValue); 
                                
                                item.Zone.SleeveInstanceId = elementId.IntegerValue;
                                
                                result.PlacedItems.Add((item.Zone, elementId));
                                result.PlacedCount++;
                            }
                            else
                            {
                                result.Failures.Add((item.Zone.Id, "Created instance was null"));
                                result.FailedCount++;
                            }
                        }
                    }

                    result.OverallSuccess = true;
                    result.ElapsedTime = DateTime.Now - startTime;
                    _logger?.Invoke($"[BulkPlacement] Successfully placed {result.PlacedCount} sleeves in {result.ElapsedTime.TotalSeconds:F2}s");
                }
            }
            catch (Exception ex)
            {
                result.OverallSuccess = false;
                result.Error = ex.Message;
                _logger?.Invoke($"[BulkPlacement] CRITICAL ERROR: {ex.Message}");
            }

            return result;
        }

        private Dictionary<string, FamilySymbol> PreActivateSymbols(Document doc, List<SleevePlacementPlanningDto> plans)
        {
            var cache = new Dictionary<string, FamilySymbol>();
            var familyNames = plans.Select(p => p.SleeveFamilyName).Distinct().Where(f => !string.IsNullOrEmpty(f));

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

        private static bool IsWallOrFraming(ClashZone zone)
        {
            if (zone?.StructuralElementType == null) return false;
            var t = zone.StructuralElementType.Trim();
            return t.IndexOf("Wall", StringComparison.OrdinalIgnoreCase) >= 0
                || t.IndexOf("Framing", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private void ApplyRotation(Document doc, FamilyInstance instance, double rotationRad)
        {
            if (Math.Abs(rotationRad) < 0.001) return;

            try
            {
                XYZ axisPoint1 = (instance.Location as LocationPoint)?.Point;
                if (axisPoint1 == null) return;
                
                XYZ axisDirection = XYZ.BasisZ; 
                XYZ axisPoint2 = axisPoint1 + axisDirection;
                Line axis = Line.CreateBound(axisPoint1, axisPoint2);
                instance.Location.Rotate(axis, rotationRad);
            }
            catch (Exception ex)
            {
                _logger?.Invoke($"[BulkPlacement] Rotation failed: {ex.Message}");
            }
        }

        private void SetWallFramingParameters(FamilyInstance instance, ClashZone zone, SleevePlacementPlanningDto plan)
        {
            try
            {
                double width = plan.TargetWidthFt;
                double height = plan.TargetHeightFt;

                if (zone != null)
                {
                    zone.SleeveWidth = width;
                    zone.SleeveHeight = height;
                }

                _parameterService.SetSleeveParameters(instance, width, height, plan.TargetDiameterFt, plan.IsCircular, zone, plan.RequiredDepthFt);
                _logger?.Invoke($"[BulkPlacement] [WALL/FRAMING] Set parameters for {instance.Id}: W={width*304.8:F0}mm, H={height*304.8:F0}mm, Circ={plan.IsCircular}");
            }
            catch (Exception ex)
            {
                _logger?.Invoke($"[BulkPlacement] [WALL/FRAMING] Failed for zone {zone?.Id}: {ex.Message}");
            }
        }

        private void SetFloorParameters(FamilyInstance instance, ClashZone zone, SleevePlacementPlanningDto plan)
        {
            try
            {
                double width = plan.TargetWidthFt;
                double height = plan.TargetHeightFt;

                if (zone != null)
                {
                    zone.SleeveWidth = width;
                    zone.SleeveHeight = height;
                }

                _parameterService.SetSleeveParameters(instance, width, height, plan.TargetDiameterFt, plan.IsCircular, zone, plan.RequiredDepthFt);
                _logger?.Invoke($"[BulkPlacement] [FLOOR] Set parameters for {instance.Id}: W={width*304.8:F0}mm, H={height*304.8:F0}mm, Circ={plan.IsCircular}");
            }
            catch (Exception ex)
            {
                _logger?.Invoke($"[BulkPlacement] [FLOOR] Failed for zone {zone?.Id}: {ex.Message}");
            }
        }

        /// <summary>
        /// Updates ClashZone objects with actual geometry from placed elements.
        /// Critical for ensuring clustering uses "As-Placed" dimensions.
        /// </summary>
        public void UpdateZonesFromElements(Document doc, List<(ClashZone Zone, ElementId Id)> items)
        {
            if (items == null || items.Count == 0) return;

            foreach (var item in items)
            {
                var zone = item.Zone;
                var elementId = item.Id;

                try
                {
                    var element = doc.GetElement(elementId);
                    if (element == null) continue;

                    UpdateZoneGeometry(zone, element);
                    
                    _logger?.Invoke($"[BulkPlacement] Updated Zone {zone.Id} from Element {element.Id}: W={zone.SleeveWidth:F3}, H={zone.SleeveHeight:F3}, Pt={zone.SleevePlacementPoint}");
                }
                catch (Exception ex)
                {
                    _logger?.Invoke($"[BulkPlacement] Failed to update zone {zone.Id} from element: {ex.Message}");
                }
            }
        }

        /// <summary>Updates zone from Revit element only. Bounding box and corners: from Revit, never calculated. See REVIT_GEOMETRY_RULES.md.</summary>
        private void UpdateZoneGeometry(ClashZone zone, Element element)
        {
            var bbox = element.get_BoundingBox(null);
            if (bbox != null)
            {
                var center = (bbox.Min + bbox.Max) * 0.5;
                
                zone.SleevePlacementPointX = center.X;
                zone.SleevePlacementPointY = center.Y;
                zone.SleevePlacementPointZ = center.Z;
                
                zone.SleevePlacementActiveX = center.X;
                zone.SleevePlacementActiveY = center.Y;
                zone.SleevePlacementActiveZ = center.Z;
                
                zone.SleevePlacementPoint = center;
                zone.SleevePlacementPointActiveDocument = center;

                double width = bbox.Max.X - bbox.Min.X;
                double height = bbox.Max.Y - bbox.Min.Y; 
                
                if (zone.SleeveWidth <= 0) zone.SleeveWidth = width;
                if (zone.SleeveHeight <= 0) zone.SleeveHeight = height;

                zone.BoundingBoxMinX = bbox.Min.X;
                zone.BoundingBoxMinY = bbox.Min.Y;
                zone.BoundingBoxMinZ = bbox.Min.Z;
                zone.BoundingBoxMaxX = bbox.Max.X;
                zone.BoundingBoxMaxY = bbox.Max.Y;
                zone.BoundingBoxMaxZ = bbox.Max.Z;

                if (element is FamilyInstance fi)
                {
                    Transform t = fi.GetTransform();
                    double rotation = t.BasisX.AngleTo(XYZ.BasisX);
                    if (t.BasisX.Y < 0) rotation = -rotation;

                    zone.CalculatedRotation = rotation;
                    zone.MepRotationCos = Math.Cos(rotation);
                    zone.MepRotationSin = Math.Sin(rotation);

                    // Rotated bounding box: from Revit bbox in element-local coordinates (no calculation from dimensions).
                    Transform ti = t.Inverse;
                    var corners = new[]
                    {
                        ti.OfPoint(bbox.Min),
                        ti.OfPoint(new XYZ(bbox.Max.X, bbox.Min.Y, bbox.Min.Z)),
                        ti.OfPoint(new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Min.Z)),
                        ti.OfPoint(new XYZ(bbox.Min.X, bbox.Max.Y, bbox.Min.Z)),
                        ti.OfPoint(new XYZ(bbox.Min.X, bbox.Min.Y, bbox.Max.Z)),
                        ti.OfPoint(new XYZ(bbox.Max.X, bbox.Min.Y, bbox.Max.Z)),
                        ti.OfPoint(bbox.Max),
                        ti.OfPoint(new XYZ(bbox.Min.X, bbox.Max.Y, bbox.Max.Z))
                    };
                    double lx0 = corners[0].X, lx1 = corners[0].X, ly0 = corners[0].Y, ly1 = corners[0].Y, lz0 = corners[0].Z, lz1 = corners[0].Z;
                    for (int i = 1; i < corners.Length; i++)
                    {
                        var c = corners[i];
                        if (c.X < lx0) lx0 = c.X; if (c.X > lx1) lx1 = c.X;
                        if (c.Y < ly0) ly0 = c.Y; if (c.Y > ly1) ly1 = c.Y;
                        if (c.Z < lz0) lz0 = c.Z; if (c.Z > lz1) lz1 = c.Z;
                    }
                    zone.RotatedBoundingBoxMinX = lx0;
                    zone.RotatedBoundingBoxMinY = ly0;
                    zone.RotatedBoundingBoxMinZ = lz0;
                    zone.RotatedBoundingBoxMaxX = lx1;
                    zone.RotatedBoundingBoxMaxY = ly1;
                    zone.RotatedBoundingBoxMaxZ = lz1;

                    // Sleeve corners: do NOT calculate here. Only persist from Revit (BatchSleeveCornerExtractor after placement).
                }
                else
                {
                    // Non-FamilyInstance: still use Revit bbox relative to placement point.
                    zone.RotatedBoundingBoxMinX = bbox.Min.X - zone.SleevePlacementPointX;
                    zone.RotatedBoundingBoxMinY = bbox.Min.Y - zone.SleevePlacementPointY;
                    zone.RotatedBoundingBoxMinZ = bbox.Min.Z - zone.SleevePlacementPointZ;
                    zone.RotatedBoundingBoxMaxX = bbox.Max.X - zone.SleevePlacementPointX;
                    zone.RotatedBoundingBoxMaxY = bbox.Max.Y - zone.SleevePlacementPointY;
                    zone.RotatedBoundingBoxMaxZ = bbox.Max.Z - zone.SleevePlacementPointZ;
                }

                if (zone.SleeveInstanceId == 0) zone.SleeveInstanceId = element.Id.IntegerValue;
            }
        }
    }
}
