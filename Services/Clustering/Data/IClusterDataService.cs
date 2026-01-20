using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using System.Collections.Generic;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Data
{
    /// <summary>
    /// Service interface for loading and caching clash zone data for clustering operations.
    /// Handles database-first data access with XML fallback for backward compatibility.
    /// </summary>
    public interface IClusterDataService
    {
        /// <summary>
        /// Load clash zones from database (primary) or XML files (fallback).
        /// Returns zones with valid SleeveInstanceId for clustering.
        /// </summary>
        /// <param name="xmlFilePath">Path to XML file (optional, used for fallback)</param>
        /// <param name="targetCategory">Target MEP category to filter (Ducts, Pipes, etc.)</param>
        /// <param name="doc">Revit document</param>
        /// <returns>List of clash zones ready for clustering</returns>
        List<ClashZone> LoadClashZonesFromRegularXml(string xmlFilePath, string targetCategory, Document doc);

        /// <summary>
        /// Load clash zone cache for cleanup operations.
        /// Populates internal cache with clash zones for fast lookups during clustering.
        /// </summary>
        /// <param name="xmlFilePath">Path to XML file (optional, used for fallback)</param>
        /// <param name="targetCategory">Target MEP category to filter</param>
        /// <param name="doc">Revit document</param>
        /// <param name="filterName">Filter name for logging</param>
        void LoadClashZoneCacheForCleanup(string xmlFilePath, string targetCategory, Document doc, string filterName);

        /// <summary>
        /// Populate cache from already-loaded clash zones.
        /// Optimization to avoid duplicate XML/database loading.
        /// </summary>
        /// <param name="clashZones">Pre-loaded clash zones</param>
        /// <param name="targetCategory">Target MEP category to filter (optional)</param>
        void LoadClashZoneCacheFromLoadedClashZones(List<ClashZone> clashZones, string targetCategory = null);

        /// <summary>
        /// Get clash zone from cache by MEP element ID.
        /// Returns null if not found.
        /// </summary>
        /// <param name="mepElementId">MEP element ID to look up</param>
        /// <returns>ClashZone or null</returns>
        ClashZone GetClashZoneFromCache(long mepElementId);

        /// <summary>
        /// Get clash zone from cache by SleeveInstanceId.
        /// Returns null if not found.
        /// </summary>
        /// <param name="sleeveInstanceId">Sleeve Instance ID to look up</param>
        /// <returns>ClashZone or null</returns>
        ClashZone GetClashZoneBySleeveInstanceId(int sleeveInstanceId);

        /// <summary>
        /// Clear the clash zone cache.
        /// </summary>
        void ClearCache();

        /// <summary>
        /// Get current cache count.
        /// </summary>
        int CacheCount { get; }

        /// <summary>
        /// Update pre-calculated 4 corner coordinates for Cluster Sleeves (Phase 3 Persistence).
        /// </summary>
        public void UpdateClusterSleeveCorners(int clusterInstanceId,
            double corner1X, double corner1Y, double corner1Z,
            double corner2X, double corner2Y, double corner2Z,
            double corner3X, double corner3Y, double corner3Z,
            double corner4X, double corner4Y, double corner4Z);

        /// <summary>
        /// Get the placement point for a cluster sleeve from the database.
        /// Returns null if not found.
        /// </summary>
        /// <param name="clusterInstanceId">Cluster sleeve instance ID</param>
        /// <returns>Placement point (XYZ) or null</returns>
        XYZ GetClusterPlacement(int clusterInstanceId);
    }
}
