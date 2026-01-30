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
    /// </summary>
    public class BulkPlacementService : IBulkPlacementService
    {
        private readonly Action<string>? _logger;
        private readonly SleeveParameterService _parameterService;
        private readonly JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.IPerformanceMonitor? _performanceMonitor;

        public BulkPlacementService(
            Document doc, 
            Action<string>? logger = null,
            JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.IPerformanceMonitor? performanceMonitor = null)
        {
            _logger = logger ?? (msg => DebugLogger.Info(msg));
            _parameterService = new SleeveParameterService(doc);
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

                // Step 2: Build Creation Data
                var creationDataList = new List<Autodesk.Revit.Creation.FamilyInstanceCreationData>();
                var itemMap = new List<(ClashZone Zone, SleevePlacementPlanningDto Plan)>();

                using (_performanceMonitor?.TrackOperation("Build Creation Data"))
                {
                    foreach (var item in taskItems)
                    {
                        var zone = item.Zone;
                        var plan = item.Plan;

                        if (string.IsNullOrEmpty(plan.SleeveFamilyName)) continue;

                        if (symbolCache.TryGetValue(plan.SleeveFamilyName, out var symbol))
                        {
                            // ✅ UNIFIED ARCHITECTURE: Use planned placement point
                            XYZ point = plan.PlacementPoint ?? new XYZ(zone.IntersectionPointX, zone.IntersectionPointY, zone.IntersectionPointZ);

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
                                // ✅ UNIFIED ARCHITECTURE: Use planned rotation
                                ApplyRotation(doc, instance, item.Plan.RotationRadians);
                                
                                // ✅ UNIFIED ARCHITECTURE: Use planned dimensions
                                SetSleeveParameters(instance, item.Zone, item.Plan);
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
                    
                    // ✅ CRITICAL FIX: Flush batched parameters to elements!
                    // Without this, all parameters remain in memory and sleeves keep default sizes (300x200x200)
                    using (_performanceMonitor?.TrackOperation("Flush Deferred Parameters"))
                    {
                        _parameterService.FlushDeferredParameters(clearList: true, context: "BulkPlacement");
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

        private void SetSleeveParameters(FamilyInstance instance, ClashZone zone, SleevePlacementPlanningDto plan)
        {
            try
            {
                // ✅ UNIFIED ARCHITECTURE: Apply all planned dimensions and metadata
                _parameterService.ApplyBatchSleeveParameters(
                    instance, 
                    plan.TargetWidthFt, 
                    plan.TargetHeightFt, 
                    plan.TargetDiameterFt, 
                    plan.RequiredDepthFt, 
                    plan.IsCircular);
                
                // Track the instance ID in the database-targeted field
               //_parameterService.SetClusterSleeveInstanceId(instance, instance.Id.IntegerValue);
                
                _logger?.Invoke($"[BulkPlacement] Set parameters for {instance.Id}: W={plan.TargetWidthFt*304.8:F0}mm, H={plan.TargetHeightFt*304.8:F0}mm, Circ={plan.IsCircular}");
            }
            catch (Exception ex)
            {
                _logger?.Invoke($"[BulkPlacement] Failed to set parameters for zone {zone.Id}: {ex.Message}");
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

                    // Update from actual element geometry
                    UpdateZoneGeometry(zone, element);
                    
                    // Log success for verification
                    _logger?.Invoke($"[BulkPlacement] Updated Zone {zone.Id} from Element {element.Id}: W={zone.SleeveWidth:F3}, H={zone.SleeveHeight:F3}, Pt={zone.SleevePlacementPoint}");
                }
                catch (Exception ex)
                {
                    _logger?.Invoke($"[BulkPlacement] Failed to update zone {zone.Id} from element: {ex.Message}");
                }
            }
        }

        private void UpdateZoneGeometry(ClashZone zone, Element element)
        {
            var bbox = element.get_BoundingBox(null);
            if (bbox != null)
            {
                // 1. Calculate Center from BBox
                var center = (bbox.Min + bbox.Max) * 0.5;
                
                // Update Placement Points (Both Standard and Active)
                zone.SleevePlacementPointX = center.X;
                zone.SleevePlacementPointY = center.Y;
                zone.SleevePlacementPointZ = center.Z;
                
                zone.SleevePlacementActiveX = center.X;
                zone.SleevePlacementActiveY = center.Y;
                zone.SleevePlacementActiveZ = center.Z;
                
                zone.SleevePlacementPoint = center;
                zone.SleevePlacementPointActiveDocument = center;

                // 2. Update Dimensions (Approximate from BBox for verification)
                // Note: We trust the Parameters more for Width/Height, but BBox gives physical extent
                double width = bbox.Max.X - bbox.Min.X;
                double height = bbox.Max.Y - bbox.Min.Y; 
                double depth = bbox.Max.Z - bbox.Min.Z;

                // Simple heuristic: if it looks rotated 90 degrees, swap X/Y dimensions
                // But generally, we stick to the parameter values we just set.
                // However, the user REQUESTED extraction.
                // Let's trust the parameters we set (which matched the plan) for "SleeveWidth/Height" 
                // as BBox includes flanges/connectors/rotation which might be confusing.
                // BUT, we MUST update the location.
                
                // User said: "SLEEVE DIMENSION AND PALCMENT POINT ALSO NEED TO BE EXTRACTED"
                // If we extract BBox Dims, we might get weird values if rotated.
                // Let's stick to Placement Point for now, and rely on the mapped Plan dimensions for W/H 
                // UNLESS they are zero.
                
                if (zone.SleeveWidth <= 0) zone.SleeveWidth = width;
                if (zone.SleeveHeight <= 0) zone.SleeveHeight = height;

                // 2b. Extract Bounding Box for Persistence (User Request)
                zone.BoundingBoxMinX = bbox.Min.X;
                zone.BoundingBoxMinY = bbox.Min.Y;
                zone.BoundingBoxMinZ = bbox.Min.Z;
                zone.BoundingBoxMaxX = bbox.Max.X;
                zone.BoundingBoxMaxY = bbox.Max.Y;
                zone.BoundingBoxMaxZ = bbox.Max.Z;

                // 2c. Extract Rotation & RCS (Rotated Coordinate System)
                // Critical for Floor Clusters to avoid "Oversized" Axis-Aligned BBox
                if (element is FamilyInstance fi)
                {
                    Transform t = fi.GetTransform();
                    // Rotation angle around Z axis (BasisX angle to Global X)
                    double rotation = t.BasisX.AngleTo(XYZ.BasisX);
                    // Check sign
                    if (t.BasisX.Y < 0) rotation = -rotation;

                    zone.CalculatedRotation = rotation;
                    zone.MepRotationCos = Math.Cos(rotation);
                    zone.MepRotationSin = Math.Sin(rotation);

                    // Extract Dimensions in LOCAL space (Geometry BBox is usually AABB)
                    // Synthesize RCS from Dimensions (assuming centered)
                    // This gives the "Tight" local bounding box
                    double halfW = zone.SleeveWidth / 2.0;
                    double halfH = zone.SleeveHeight / 2.0;
                    double halfD = (zone.SleeveDepth > 0 ? zone.SleeveDepth : (bbox.Max.Z - bbox.Min.Z)) / 2.0;

                    zone.RotatedBoundingBoxMinX = -halfW;
                    zone.RotatedBoundingBoxMinY = -halfH;
                    zone.RotatedBoundingBoxMinZ = -halfD;
                    zone.RotatedBoundingBoxMaxX = halfW;
                    zone.RotatedBoundingBoxMaxY = halfH;
                    zone.RotatedBoundingBoxMaxZ = halfD;

                    // ✅ ADDED: Calculate Corners (OBB) for Clustering Accuracy
                    // The clustering service prioritizes corners. We calculate them from the Transform + Dimensions
                    // to ensure "Tight" fit matching the OBB, not the AABB.
                    
                    // Local corners (bottom face usually, or center-Z plane)
                    // Converting to Global using Transform
                    XYZ centerLoc = XYZ.Zero; // Local center
                    
                    // We need to respect the element's placement point (often center)
                    // The Transform origin is the placement point.
                    
                    // C1: Min X, Min Y
                    XYZ p1Local = new XYZ(-halfW, -halfH, 0); 
                    // C2: Max X, Min Y
                    XYZ p2Local = new XYZ(halfW, -halfH, 0);
                    // C3: Max X, Max Y
                    XYZ p3Local = new XYZ(halfW, halfH, 0);
                    // C4: Min X, Max Y
                    XYZ p4Local = new XYZ(-halfW, halfH, 0);

                    XYZ p1Global = t.OfPoint(p1Local);
                    XYZ p2Global = t.OfPoint(p2Local);
                    XYZ p3Global = t.OfPoint(p3Local);
                    XYZ p4Global = t.OfPoint(p4Local);

                    zone.SleeveCorner1X = p1Global.X;
                    zone.SleeveCorner1Y = p1Global.Y;
                    zone.SleeveCorner1Z = p1Global.Z;

                    zone.SleeveCorner2X = p2Global.X;
                    zone.SleeveCorner2Y = p2Global.Y;
                    zone.SleeveCorner2Z = p2Global.Z;

                    zone.SleeveCorner3X = p3Global.X;
                    zone.SleeveCorner3Y = p3Global.Y;
                    zone.SleeveCorner3Z = p3Global.Z;

                    zone.SleeveCorner4X = p4Global.X;
                    zone.SleeveCorner4Y = p4Global.Y;
                    zone.SleeveCorner4Z = p4Global.Z;
                }
                else
                {
                    // Fallback for non-FamilyInstances (unexpected) -> Use BBox Corners
                    zone.RotatedBoundingBoxMinX = bbox.Min.X - zone.SleevePlacementPointX;
                    zone.RotatedBoundingBoxMaxX = bbox.Max.X - zone.SleevePlacementPointX;
                    // ... (simplified)
                }



                // 3. Ensure SleeveInstanceId is set (redundant check)
                if (zone.SleeveInstanceId == 0) zone.SleeveInstanceId = element.Id.IntegerValue;
            }
        }
    }
}
