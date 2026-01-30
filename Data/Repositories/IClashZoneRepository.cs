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
        /// ✅ BATCH OPTIMIZATION: Insert or update clash zones in a single multi-category batch.
        /// Consolidates multiple transactions into one atomic operation.
        /// </summary>
        void InsertOrUpdateClashZonesBulk(IEnumerable<ClashZone> clashZones, string filterName);

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
            double rotationAngleRad, string sleeveFamilyName = null, bool markedForClusterProcess = true);

        /// <summary>
        /// Update cluster placement
        /// </summary>
        void UpdateClusterPlacement(System.Guid clashZoneId, int clusterInstanceId, double minX, double minY, double minZ,
            double maxX, double maxY, double maxZ, double? placementX = null, double? placementY = null, double? placementZ = null,
            double? rotatedMinX = null, double? rotatedMinY = null, double? rotatedMinZ = null,
            double? rotatedMaxX = null, double? rotatedMaxY = null, double? rotatedMaxZ = null,
            bool? isClustered = null, bool? markedForCluster = null, string? sleeveFamilyName = null,
            double sleeveWidth = 0, double sleeveHeight = 0, double sleeveDiameter = 0);

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
        /// Update the MEP Category for a clash zone (dump once, use many times)
        /// </summary>
        void UpdateMepCategory(System.Guid clashZoneGuid, string category);

        /// <summary>
        /// Update the Sleeve Family Name for a clash zone (persists pre-placement determination)
        /// </summary>
        void UpdateSleeveFamilyName(System.Guid clashZoneGuid, string familyName);

        /// <summary>
        /// Update the Sleeve Family Name for multiple clash zones (persists pre-placement determination)
        /// </summary>
        void UpdateSleeveFamilyNameBulk(System.Collections.Generic.IEnumerable<System.Guid> clashZoneGuids, string familyName);

        /// <summary>
        /// Get MEP Categories for a list of sleeve instance IDs
        /// Returns dictionary of SleeveInstanceId -> MepCategory
        /// </summary>
        System.Collections.Generic.Dictionary<int, string> GetMepCategoriesForSleeveIds(System.Collections.Generic.IEnumerable<int> sleeveInstanceIds);

        /// <summary>
        /// Updates the 4 corner coordinates for a cluster sleeve in the database.
        /// </summary>
        void UpdateClusterSleeveCorners(int clusterInstanceId,
            double corner1X, double corner1Y, double corner1Z,
            double corner2X, double corner2Y, double corner2Z,
            double corner3X, double corner3Y, double corner3Z,
            double corner4X, double corner4Y, double corner4Z);

        /// <summary>
        /// Log sleeve event
        /// </summary>
        void LogSleeveEvent(int clashZoneId, string eventType, string? payload = null);

        /// <summary>
        /// Retrieve MEP snapshot parameter key/value pairs for a list of sleeve element instance ids.
        /// Returns a dictionary keyed by SleeveInstanceId with a dictionary of parameter name->value (string).
        /// Empty dictionary returned if none found or input invalid. Host/cluster parameters are excluded for now.
        /// </summary>
        /// <param name="sleeveInstanceIds">Collection of Revit element ids for sleeves.</param>
        /// <returns>Dictionary<int, Dictionary<string,string>></returns>
        System.Collections.Generic.Dictionary<int, System.Collections.Generic.Dictionary<string, string>> GetSnapshotMepParametersForSleeveIds(System.Collections.Generic.IEnumerable<int> sleeveInstanceIds);

        /// <summary>
        /// Batch update IsResolvedFlag, IsClusterResolvedFlag, IsCombinedResolved, SleeveInstanceId, and ClusterInstanceId for placed sleeves.
        /// Now includes IsClusteredFlag, MarkedForClusterProcess, and AfterClusterSleeveId.
        /// </summary>
        void BatchUpdateFlags(List<(System.Guid ClashZoneId, bool IsResolved, bool IsClusterResolved, bool IsCombinedResolved, int SleeveInstanceId, int ClusterInstanceId, bool IsClusteredFlag, bool MarkedForClusterProcess, int AfterClusterSleeveId)> updates);

        /// <summary>
        /// Batch update flags including IsCurrentClashFlag.
        /// </summary>
        void BatchUpdateFlagsWithCurrentClash(List<(System.Guid ClashZoneId, int ClashZoneIntId, bool IsResolvedFlag, bool IsClusterResolvedFlag, bool IsCombinedResolved, int SleeveInstanceId, int ClusterInstanceId, bool IsCurrentClashFlag, bool IsClusteredFlag, bool? MarkedForClusterProcess, int AfterClusterSleeveId, double SleeveWidth, double SleeveHeight, double SleeveDiameter)> updates);

        /// <summary>
        /// Force Detection Mode: Reset all flags (IsResolved, IsClusterResolved) to false and clear sleeve IDs
        /// for all zones in the specified filters and categories, while preserving GUIDs.
        /// Used when ForceDetectionMode is enabled in global settings.
        /// </summary>
        int ResetAllFlagsForForceDetectionMode(List<string> filterNames, List<string> categories);

        /// <summary>
        /// Verify if sleeves marked as resolved in the DB still exist in the current Revit model.
        /// If not found, reset their flags to false and IDs to -1.
        /// </summary>
        int VerifyExistingSleevesAndResetFlags(Autodesk.Revit.DB.Document doc, List<string> filterNames, List<string> categories);

        /// <summary>
        /// Reset IsCurrentClashFlag to false for all zones in the specified filters/categories.
        /// This is called at the start of a refresh cycle to mark all existing zones as "stale" until re-detected.
        /// </summary>
        int ResetIsCurrentClashFlag(List<string> filterNames, List<string> categories);

        /// <summary>
        /// Reset IsFilterComboNew flag to 0 for a specific combo ID.
        /// </summary>
        void ResetFileComboFlag(int comboId);

        /// <summary>
        /// Retrieve ClashZone objects associated with the given Revit Sleeve Instance IDs.
        /// This checks both individual SleeveInstanceId and ClusterSleeveInstanceId.
        /// </summary>
        List<ClashZone> GetClashZonesBySleeveIds(IEnumerable<int> sleeveInstanceIds);

        /// <summary>
        /// Retrieve the ComboId and FilterId for a given ClashZone GUID.
        /// </summary>
        (int ComboId, int FilterId) GetComboAndFilterId(System.Guid clashZoneId);

        /// <summary>
        /// Retrieves a list of ClashZones by their Guids.
        /// </summary>
        List<ClashZone> GetClashZonesByGuids(IEnumerable<System.Guid> guids);

        /// <summary>
        /// Retrieves cluster sleeves associated with the given instance IDs.
        /// </summary>
        List<ClusterSleeve> GetClusterSleevesByInstanceIds(IEnumerable<int> instanceIds);

        /// <summary>
        /// Retrieves all cluster sleeves from the database.
        /// </summary>
        List<ClusterSleeve> GetAllClusterSleeves();

        /// <summary>
        /// Retrieves distinct MEP categories from the ClashZones table.
        /// </summary>


        /// <summary>
        /// Updates resolution flags for zones that are part of a combined sleeve.
        /// Sets IsCombinedResolved=true, IsResolved=false, IsClusterResolved=false, 
        /// and links to the combined sleeve ID.
        /// </summary>
        void UpdateCombinedResolutionFlags(IEnumerable<System.Guid> zoneGuids, int combinedSleeveId);

        /// <summary>
        /// Updates resolution flags for all zones belonging to the specified cluster instance IDs.
        /// </summary>
        void UpdateCombinedResolutionFlagsByClusterIds(IEnumerable<int> clusterInstanceIds, int combinedSleeveId);

        /// <summary>
        /// ✅ BATCH OPTIMIZATION: Finds existing GUIDs for a collection of MEP+Host+Point triples.
        /// Returns a dictionary mapping (MepId, HostId, RoundedPoint) -> Guid.
        /// </summary>
        System.Collections.Generic.Dictionary<(int MepId, int HostId, string PointKey), System.Guid> FindGuidsByMepHostAndPointsBulk(
            System.Collections.Generic.IEnumerable<(int MepId, int HostId, double X, double Y, double Z)> targets, 
            double tolerance = 0.001);

        /// <summary>
        /// ✅ BATCH OPTIMIZATION: Update R-tree index for multiple clash zones in one pass.
        /// </summary>
        void BulkUpdateRTreeIndex(IEnumerable<ClashZone> zones, System.Data.SQLite.SQLiteTransaction? transaction = null);



        /// <summary>
        /// ✅ PLACEMENT OPTIMIZATION: Batch update sleeve placement data in a single transaction.
        /// Replaces multiple UpdateSleevePlacement calls with one batch operation (50x faster).
        /// </summary>
        void BatchUpdateSleevePlacement(
            IEnumerable<(System.Guid ClashZoneGuid, int SleeveInstanceId, double Width, double Height, double Diameter,
                double PlacementX, double PlacementY, double PlacementZ,
                double PlacementActiveX, double PlacementActiveY, double PlacementActiveZ,
                double RotationAngleRad, string SleeveFamilyName)> updates);

        /// <summary>
        /// ✅ PLACEMENT OPTIMIZATION: Batch update cluster placement data in a single transaction.
        /// </summary>
        void BatchUpdateClusterPlacement(
            IEnumerable<(System.Guid ClashZoneGuid, int ClusterInstanceId, 
                double Width, double Height, double Diameter,
                double BoundingBoxMinX, double BoundingBoxMinY, double BoundingBoxMinZ,
                double BoundingBoxMaxX, double BoundingBoxMaxY, double BoundingBoxMaxZ,
                string SleeveFamilyName)> updates);

        /// <summary>
        /// ✅ PLACEMENT OPTIMIZATION: Batch update SleeveInstanceId for multiple zones in a single transaction.
        /// </summary>
        void BatchUpdateSleeveInstanceIds(IEnumerable<(System.Guid ClashZoneId, int SleeveInstanceId)> updates);



        /// <summary>
        /// ✅ DEBUGTOOL: Reset IsResolved and IsClusterResolved flags to false for all zones 
        /// whose Center point falls within the given Section Box (World Coordinates).
        /// Returns the number of rows affected.
        /// </summary>
        int ResetResolvedFlagsInSectionBox(Autodesk.Revit.DB.BoundingBoxXYZ sectionBox);

        /// <summary>
        /// ✅ DEBUG: Get flag statistics for all zones in DB
        /// </summary>
        (int Total, int IsCurrentClashSet, int ReadyForPlacementSet, int IsResolvedSet) GetFlagStatistics();

        /// <summary>
        /// Update calculated sleeve data (Phase 1 result)
        /// </summary>
        void UpdateCalculatedSleeveData(System.Guid clashZoneGuid,
            double width, double height, double depth,
            double rotation, string familyName,
            string status, string batchId);

        /// <summary>
        /// Update sleeve corners in bulk
        /// </summary>
        void BatchUpdateSleeveCorners(IEnumerable<(Guid Guid, double c1x, double c1y, double c1z, double c2x, double c2y, double c2z, double c3x, double c3y, double c3z, double c4x, double c4y, double c4z)> updates);

        /// <summary>
        /// Get distinct categories present in the database
        /// </summary>
        List<string> GetDistinctCategories();

        /// <summary>
        /// Retrieve all clash zones from the database.
        /// </summary>
        List<ClashZone> GetAllClashZones();

        /// <summary>
        /// Update a single clash zone.
        /// </summary>
        void Update(ClashZone zone);
        
        /// <summary>
        /// Retrieves all clash zones that have been placed (SleeveInstanceId > 0).
        /// </summary>
        List<ClashZone> GetPlacedClashZones();
        
        /// <summary>
        /// ✅ SELF-HEALING: Try to get SleeveInstanceId from SleeveSnapshots table if missing in ClashZones.
        /// Used for recovering lost IDs during cleanup/swap operations.
        /// </summary>
        int TryGetSleeveInstanceIdFromSnapshot(System.Guid clashZoneGuid);
        
        /// <summary>
        /// Save snapshots for placed sleeves to the SleeveSnapshots table.
        /// Critical for Self-Healing Persistence.
        /// </summary>
        void SaveSleeveSnapshotsForPlacedSleeves(int filterId, List<ClashZone> placedZones);
    }
}
