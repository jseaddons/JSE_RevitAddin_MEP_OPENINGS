using System;

namespace JSE_RevitAddin_MEP_OPENINGS.Data.Entities
{
    /// <summary>
    /// SQLite entity for ClashZones table
    /// Maps to existing ClashZone model
    /// </summary>
    public class SleeveClashZone
    {
        public int ClashZoneId { get; set; }
        public int ComboId { get; set; }
        public long MepElementId { get; set; }
        public long HostElementId { get; set; }
        public double IntersectionX { get; set; }
        public double IntersectionY { get; set; }
        public double IntersectionZ { get; set; }
        public int SleeveState { get; set; } // 0=Unprocessed, 1=IndividualPlaced, 2=ClusterPlaced, etc.
        public int? SleeveInstanceId { get; set; }
        public int? ClusterInstanceId { get; set; }
        public double? SleeveWidth { get; set; }
        public double? SleeveHeight { get; set; }
        public double? SleeveDiameter { get; set; }
        public double? SleevePlacementX { get; set; }
        public double? SleevePlacementY { get; set; }
        public double? SleevePlacementZ { get; set; }
        public double? BoundingBoxMinX { get; set; }
        public double? BoundingBoxMinY { get; set; }
        public double? BoundingBoxMinZ { get; set; }
        public double? BoundingBoxMaxX { get; set; }
        public double? BoundingBoxMaxY { get; set; }
        public double? BoundingBoxMaxZ { get; set; }
        public string? PlacementSource { get; set; }
        public DateTime UpdatedAt { get; set; }
        public string? ClashZoneGuid { get; set; }
        public string? MepCategory { get; set; }
        public string? StructuralType { get; set; }
        public string? HostOrientation { get; set; }
        public string? MepOrientationDirection { get; set; }
        public double? MepOrientationX { get; set; }
        public double? MepOrientationY { get; set; }
        public double? MepOrientationZ { get; set; }
        public double? MepRotationAngleRad { get; set; }
        public double? MepRotationAngleDeg { get; set; }
        // ✅ ROTATION MATRIX: Pre-calculated cos/sin for "dump once use many times" principle
        public double? MepRotationCos { get; set; }
        public double? MepRotationSin { get; set; }
        public double? MepAngleToXRad { get; set; }
        public double? MepAngleToXDeg { get; set; }
        public double? MepAngleToYRad { get; set; }
        public double? MepAngleToYDeg { get; set; }
        public double? MepWidth { get; set; }
        public double? MepHeight { get; set; }
        public string? SleeveFamilyName { get; set; }
        public double? SleevePlacementActiveX { get; set; }
        public double? SleevePlacementActiveY { get; set; }
        public double? SleevePlacementActiveZ { get; set; }
        public string? SourceDocKey { get; set; }
        public string? HostDocKey { get; set; }
        public string? MepElementUniqueId { get; set; }
        public bool IsResolvedFlag { get; set; }
        public bool IsClusterResolvedFlag { get; set; }
        public bool? IsClusteredFlag { get; set; }
        public bool? MarkedForClusterProcess { get; set; }
        public int AfterClusterSleeveId { get; set; }
        public bool HasDamperNearbyFlag { get; set; }
        public bool IsCurrentClashFlag { get; set; }
        // ✅ SLEEVE CORNERS: Pre-calculated 4 corner coordinates in world space (for clustering optimization)
        public double? SleeveCorner1X { get; set; }
        public double? SleeveCorner1Y { get; set; }
        public double? SleeveCorner1Z { get; set; }
        public double? SleeveCorner2X { get; set; }
        public double? SleeveCorner2Y { get; set; }
        public double? SleeveCorner2Z { get; set; }
        public double? SleeveCorner3X { get; set; }
        public double? SleeveCorner3Y { get; set; }
        public double? SleeveCorner3Z { get; set; }
        public double? SleeveCorner4X { get; set; }
        public double? SleeveCorner4Y { get; set; }
        public double? SleeveCorner4Z { get; set; }
    }
}

