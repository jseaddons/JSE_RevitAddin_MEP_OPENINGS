using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor
{
    /// <summary>
    /// Team D: Persistence service abstraction for sleeve placement data.
    /// Encapsulates DB updates and provides transaction boundaries.
    /// All persistence operations route through this abstraction.
    /// </summary>
    public interface ISleevePersistenceService
    {
        /// <summary>
        /// Persists placement data for a collection of placed sleeves.
        /// Updates database with sleeve instance IDs, dimensions, and placement coordinates.
        /// </summary>
        /// <param name="placedSleeves">Collection of (FamilyInstance, ClashZone) tuples representing placed sleeves</param>
        /// <param name="filterName">Filter name for context</param>
        /// <param name="category">MEP category for context</param>
        /// <returns>Number of sleeves successfully persisted</returns>
        Task<int> PersistPlacementAsync(
            IEnumerable<(FamilyInstance sleeve, ClashZone zone)> placedSleeves,
            string filterName,
            string category);
        
        /// <summary>
        /// Updates a single sleeve instance in the database.
        /// Used for individual updates after placement.
        /// </summary>
        /// <param name="zone">ClashZone with updated sleeve data</param>
        /// <param name="instance">FamilyInstance that was placed</param>
        /// <returns>True if update succeeded</returns>
        Task<bool> UpdateInstanceAsync(ClashZone zone, FamilyInstance instance);
        
        /// <summary>
        /// Updates sleeve bounding box coordinates in the database.
        /// </summary>
        /// <param name="zone">ClashZone with bounding box data</param>
        /// <param name="isClusterSleeve">Whether this is a cluster sleeve</param>
        /// <returns>True if update succeeded</returns>
        Task<bool> UpdateBoundingBoxAsync(ClashZone zone, bool isClusterSleeve = false);
        
        /// <summary>
        /// Saves sleeve snapshots for placed sleeves (for parameter transfer).
        /// </summary>
        /// <param name="filterId">Filter ID</param>
        /// <param name="placedZones">List of placed clash zones</param>
        /// <returns>Number of snapshots saved</returns>
        Task<int> SaveSnapshotsAsync(int filterId, List<ClashZone> placedZones);
    }
}

