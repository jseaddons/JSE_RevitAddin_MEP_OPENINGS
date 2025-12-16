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

            // 1. Convert Elements to Dynamic Objects for Core Services
            // The Core Services expect a List<dynamic> where each item has SleeveInstanceId property.
            var clusterData = new List<dynamic>();
            foreach (var sleeve in sleeves)
            {
                clusterData.Add(new { SleeveInstanceId = sleeve.Id.IntegerValue });
            }

            // 2. Delegate: Calculate Rotation using "Golden" Logic
            // This handles X-Wall vs Y-Wall vs Circular logic automatically.
            double rotationAngle = _rotationService.DetermineRotationAngle(clusterData);

            // 3. Delegate: Calculate Bounding Box using "Golden" Logic
            // This uses the watertight World Space analysis.
            var result = _bboxCalculator.Calculate(
                clusterData, 
                sleeves.OfType<FamilyInstance>().ToList(), 
                rotationAngle
            );

            // 4. Return Result
            return new ManualJoinResult
            {
                WidthFeet = result.Width,
                HeightFeet = result.Height,
                DepthFeet = result.Depth,
                RotationAngleRad = rotationAngle,
                CenterPoint = result.Midpoint
            };
        }
    }
}
