using System.Collections.Generic;
using System.Threading.Tasks;
using JSE_RevitAddin_MEP_OPENINGS.Data.Entities;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Data.Repositories
{
    /// <summary>
    /// Repository interface for ClashZone persistence
    /// Abstracts SQLite operations from business logic
    /// </summary>
    public interface IClashZoneRepository
    {
        /// <summary>
        /// Insert or update clash zones (dual-write mode: writes to both XML and SQLite)
        /// </summary>
        void InsertOrUpdateClashZones(IEnumerable<ClashZone> clashZones, string filterName, string category);

        /// <summary>
        /// Get clash zones by filter and category
        /// </summary>
        List<ClashZone> GetClashZonesByFilter(string filterName, string category, bool unresolvedOnly = false, bool readyForPlacementOnly = false);

        /// <summary>
        /// Get clash zones by category (all filters)
        /// </summary>
        List<ClashZone> GetClashZonesByCategory(string category);

        /// <summary>
        /// Update sleeve state and dimensions
        /// ✅ Also saves Active document coordinates (where sleeve is actually placed)
        /// ✅ Also saves rotation angle (for corner calculations)
        /// </summary>
        void UpdateSleevePlacement(System.Guid clashZoneGuid, int sleeveInstanceId, double width, double height, double diameter, 
            double placementX, double placementY, double placementZ,
            double placementActiveX, double placementActiveY, double placementActiveZ,
            double rotationAngleRad);

        /// <summary>
        /// Update cluster placement
        /// </summary>
        void UpdateClusterPlacement(int clashZoneId, int clusterInstanceId, double minX, double minY, double minZ,
            double maxX, double maxY, double maxZ, double? placementX = null, double? placementY = null, double? placementZ = null,
            double? rotatedMinX = null, double? rotatedMinY = null, double? rotatedMinZ = null,
            double? rotatedMaxX = null, double? rotatedMaxY = null, double? rotatedMaxZ = null,
            bool? isClustered = null, bool? markedForCluster = null);

        /// <summary>
        /// Update rotated bounding box coordinates for a rotated individual sleeve
        /// </summary>
        void UpdateRotatedBoundingBoxes(System.Guid clashZoneGuid, double rotatedMinX, double rotatedMinY, double rotatedMinZ,
            double rotatedMaxX, double rotatedMaxY, double rotatedMaxZ);

        /// <summary>
        /// ✅ SLEEVE CORNERS: Update pre-calculated 4 corner coordinates in world space
        /// Calculated once during individual sleeve placement, stored for reuse during clustering
        /// </summary>
        void UpdateSleeveCorners(System.Guid clashZoneGuid, 
            double corner1X, double corner1Y, double corner1Z,
            double corner2X, double corner2Y, double corner2Z,
            double corner3X, double corner3Y, double corner3Z,
            double corner4X, double corner4Y, double corner4Z);

        /// <summary>
        /// Log sleeve event
        /// </summary>
        void LogSleeveEvent(int clashZoneId, string eventType, string? payload = null);
    }
}

