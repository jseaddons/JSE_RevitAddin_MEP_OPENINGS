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
        List<ClashZone> GetClashZonesByFilter(string filterName, string category, bool unresolvedOnly = false);

        /// <summary>
        /// Get clash zones by category (all filters)
        /// </summary>
        List<ClashZone> GetClashZonesByCategory(string category);

        /// <summary>
        /// Update sleeve state and dimensions
        /// </summary>
        void UpdateSleevePlacement(System.Guid clashZoneGuid, int sleeveInstanceId, double width, double height, double diameter, 
            double placementX, double placementY, double placementZ);

        /// <summary>
        /// Update cluster placement
        /// </summary>
        void UpdateClusterPlacement(int clashZoneId, int clusterInstanceId, double minX, double minY, double minZ,
            double maxX, double maxY, double maxZ);

        /// <summary>
        /// Log sleeve event
        /// </summary>
        void LogSleeveEvent(int clashZoneId, string eventType, string? payload = null);
    }
}

