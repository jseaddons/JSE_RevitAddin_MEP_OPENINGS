using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.BoundingBox;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Rotation;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Geometry;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Combined
{
    public interface IManualClusterCalculationAdapter
    {
        ManualJoinResult Calculate(List<Element> sleeves);
    }

    public class ManualJoinResult
    {
        public double WidthFeet { get; set; }
        public double HeightFeet { get; set; }
        public double DepthFeet { get; set; }
        public double RotationAngleRad { get; set; }
        public XYZ CenterPoint { get; set; }
    }

    /// <summary>
    /// Adapter to bridge Manual Join UI with the "Golden" Auto-Cluster Core Services.
    /// This ensures Manual Join produces mathematically identical results to Auto-Clustering.
    /// </summary>
    public class ManualClusterCalculationAdapter : IManualClusterCalculationAdapter
    {
        private readonly IClusterRotationService _rotationService;
        private readonly IBoundingBoxCalculator _bboxCalculator;
        private readonly ClashZoneRepository _repo;

        public ManualClusterCalculationAdapter(
            IClusterRotationService rotationService,
            IBoundingBoxCalculator bboxCalculator,
            ClashZoneRepository repo)
        {
            _rotationService = rotationService ?? throw new ArgumentNullException(nameof(rotationService));
            _bboxCalculator = bboxCalculator ?? throw new ArgumentNullException(nameof(bboxCalculator));
            _repo = repo ?? throw new ArgumentNullException(nameof(repo));
        }

        public ManualJoinResult Calculate(List<Element> sleeves)
        {
            if (sleeves == null || sleeves.Count == 0)
                throw new ArgumentException("Sleeves list cannot be empty");

            var sleeveIds = sleeves.Select(s => s.Id.IntegerValue).ToList();
            
            JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Info(
                $"[ManualClusterCalculationAdapter] Fetching corner data for {sleeveIds.Count} sleeves from DB");
            
            // ✅ Get corners from BOTH Individual AND Cluster sleeves
            var allCorners = new List<XYZ>();
            double? firstRotation = null;
            
            // 1. Try Individual Sleeves (ClashZones table)
            // Use naming convention from Model (SleeveCorner1X, not Corner1X)
            var clashZones = _repo.GetClashZonesBySleeveIds(sleeveIds);
            foreach (var cz in clashZones)
            {
                if (cz != null && 
                    cz.SleeveCorner1X.HasValue && cz.SleeveCorner1Y.HasValue && cz.SleeveCorner1Z.HasValue &&
                    cz.SleeveCorner2X.HasValue && cz.SleeveCorner2Y.HasValue && cz.SleeveCorner2Z.HasValue &&
                    cz.SleeveCorner3X.HasValue && cz.SleeveCorner3Y.HasValue && cz.SleeveCorner3Z.HasValue &&
                    cz.SleeveCorner4X.HasValue && cz.SleeveCorner4Y.HasValue && cz.SleeveCorner4Z.HasValue)
                {
                    allCorners.Add(new XYZ(cz.SleeveCorner1X.Value, cz.SleeveCorner1Y.Value, cz.SleeveCorner1Z.Value));
                    allCorners.Add(new XYZ(cz.SleeveCorner2X.Value, cz.SleeveCorner2Y.Value, cz.SleeveCorner2Z.Value));
                    allCorners.Add(new XYZ(cz.SleeveCorner3X.Value, cz.SleeveCorner3Y.Value, cz.SleeveCorner3Z.Value));
                    allCorners.Add(new XYZ(cz.SleeveCorner4X.Value, cz.SleeveCorner4Y.Value, cz.SleeveCorner4Z.Value));
                    
                    if (!firstRotation.HasValue && Math.Abs(cz.MepElementRotationAngle) > 0.0001)
                        firstRotation = cz.MepElementRotationAngle * 180.0 / Math.PI;

                    JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Info($"[ManualClusterCalculationAdapter] ✅ Read DB Corners for Sleeve {cz.SleeveInstanceId}: ({cz.SleeveCorner1X:F2},{cz.SleeveCorner1Y:F2})...");
                }
            }
            
            // 2. Try Cluster Sleeves (ClusterSleeves table)
            var clusterSleeves = _repo.GetClusterSleevesByInstanceIds(sleeveIds);
            foreach (var cluster in clusterSleeves)
            {
                // ClusterSleeves model uses Corner1X... naming (verified in repo map method)
                if (cluster.Corner1X.HasValue && cluster.Corner1Y.HasValue && cluster.Corner1Z.HasValue &&
                    cluster.Corner2X.HasValue && cluster.Corner2Y.HasValue && cluster.Corner2Z.HasValue &&
                    cluster.Corner3X.HasValue && cluster.Corner3Y.HasValue && cluster.Corner3Z.HasValue &&
                    cluster.Corner4X.HasValue && cluster.Corner4Y.HasValue && cluster.Corner4Z.HasValue)
                {
                    allCorners.Add(new XYZ(cluster.Corner1X.Value, cluster.Corner1Y.Value, cluster.Corner1Z.Value));
                    allCorners.Add(new XYZ(cluster.Corner2X.Value, cluster.Corner2Y.Value, cluster.Corner2Z.Value));
                    allCorners.Add(new XYZ(cluster.Corner3X.Value, cluster.Corner3Y.Value, cluster.Corner3Z.Value));
                    allCorners.Add(new XYZ(cluster.Corner4X.Value, cluster.Corner4Y.Value, cluster.Corner4Z.Value));
                    
                    if (!firstRotation.HasValue && cluster.RotationAngleDeg.HasValue)
                        firstRotation = cluster.RotationAngleDeg.Value; // Cluster table stores Deg

                    JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Info($"[ManualClusterCalculationAdapter] ✅ Read DB Corners for Cluster {cluster.ClusterInstanceId}: ({cluster.Corner1X:F2},{cluster.Corner1Y:F2})...");
                }
            }
            
            if (allCorners.Count == 0)
            {
                 // Fallback to calculator if NO corners found in DB
                 JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Warning("[ManualClusterCalculationAdapter] No Corners in DB. Using Fallback Calculator.");
                 
                 // Need rotation service for fallback
                 var clusterData = new List<JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Data.ClusteringSleeveDto>();
                 foreach(var s in sleeves) 
                 {
                     var cz = clashZones.FirstOrDefault(c => c.SleeveInstanceId == s.Id.IntegerValue);
                     clusterData.Add(new JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Data.ClusteringSleeveDto 
                     { 
                         SleeveInstanceId = s.Id.IntegerValue, 
                         ClashZone = cz 
                     });
                 }
                 
                 double fbRot = _rotationService.DetermineRotationAngle(clusterData);
                 
                 
                 // Adapter: BoundingBoxCalculator accepts List<ClusteringSleeveDto>
                 var bbox = _bboxCalculator.Calculate(clusterData, sleeves.OfType<FamilyInstance>().ToList(), fbRot);
                 
                 return new ManualJoinResult
                 {
                    WidthFeet = bbox.Width,
                    HeightFeet = bbox.Height,
                    DepthFeet = bbox.Depth,
                    RotationAngleRad = fbRot,
                    CenterPoint = bbox.Midpoint
                 };
            }
            
            // ✅ Calculate extremities from corners
            double minX = allCorners.Min(c => c.X);
            double maxX = allCorners.Max(c => c.X);
            double minY = allCorners.Min(c => c.Y);
            double maxY = allCorners.Max(c => c.Y);
            double minZ = allCorners.Min(c => c.Z);
            double maxZ = allCorners.Max(c => c.Z);
            
            double xRange = maxX - minX;
            double yRange = maxY - minY;
            double zRange = maxZ - minZ;
            
            // Determine Orientation from the first valid ClashZone or ClusterSleeve
            string orientation = "X"; // Default
            string hostType = "Wall"; // Default
            
            var refZone = clashZones.FirstOrDefault(c => !string.IsNullOrEmpty(c.HostOrientation));
            if (refZone != null)
            {
                orientation = refZone.HostOrientation;
                hostType = refZone.StructuralElementType;
            }
            else
            {
                // Try clusters
                var refCluster = clusterSleeves.FirstOrDefault(c => !string.IsNullOrEmpty(c.HostOrientation)); // Ensure ClusterSleeve model has this
                if (refCluster != null)
                {
                    orientation = refCluster.HostOrientation;
                    hostType = refCluster.HostType;
                }
            }

            double width, height, depth;
            
            if (string.Equals(orientation, "Y", StringComparison.OrdinalIgnoreCase))
            {
                // Y-Wall
                width = yRange;  // Length along wall
                height = zRange; // Vertical height
                depth = xRange;  // Wall thickness
                JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Info($"[ManualAdapter] Host Orientation: Y (Width=Y, Depth=X)");
            }
            else if (string.Equals(orientation, "Z", StringComparison.OrdinalIgnoreCase) || 
                     string.Equals(hostType, "Floor", StringComparison.OrdinalIgnoreCase))
            {
                // Floor or Vertical Host
                width = xRange;
                height = yRange;
                depth = zRange; // Floor thickness
                JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Info($"[ManualAdapter] Host Orientation: Floor/Z (Width=X, Depth=Z)");
            }
            else
            {
                // Default / X-Wall
                width = xRange;
                height = zRange; // Vertical height
                depth = yRange;  // Wall thickness
                 JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Info($"[ManualAdapter] Host Orientation: X (Width=X, Depth=Y)");
            }
            
            XYZ center = new XYZ(
                (minX + maxX) / 2.0,
                (minY + maxY) / 2.0,
                (minZ + maxZ) / 2.0);
            
            // Determine rotation
            double rotationRad = 0.0;
            
            // Let's re-fetch rotation for consistency without complex mixing
            // The user wanted simple logic.
            // Let's rely on the first loop finding MepElementRotationAngle.
            var firstCZ = clashZones.FirstOrDefault(c => Math.Abs(c.MepElementRotationAngle) > 0.0001);
            if (firstCZ != null) rotationRad = firstCZ.MepElementRotationAngle; // Already Rads
            else 
            {
                // check clusters
                var firstCl = clusterSleeves.FirstOrDefault(c => c.RotationAngleDeg.HasValue);
                if (firstCl != null) rotationRad = firstCl.RotationAngleDeg.Value * Math.PI / 180.0;
            }

            JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Info(
                $"[ManualClusterCalculationAdapter] ✅ Calculated from {allCorners.Count} corners:");
            JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Info(
                $"  Width: {width * 304.8:F1} mm, Height: {height * 304.8:F1} mm, Depth: {depth * 304.8:F1} mm");
            
            return new ManualJoinResult
            {
                WidthFeet = width,
                HeightFeet = height,
                DepthFeet = depth,
                RotationAngleRad = rotationRad,
                CenterPoint = center
            };
        }
    }
}
