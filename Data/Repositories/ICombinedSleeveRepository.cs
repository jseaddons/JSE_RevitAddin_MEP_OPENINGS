using System.Collections.Generic;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Data.Repositories
{
    /// <summary>
    /// Repository interface for cross-category combined sleeve operations.
    /// Provides CRUD operations for combined sleeves and their constituents.
    /// </summary>
    public interface ICombinedSleeveRepository
    {
        // ============================================================================
        // CREATE OPERATIONS
        // ============================================================================
        
        /// <summary>
        /// Saves a combined sleeve to the database (including constituents).
        /// Returns the generated CombinedSleeveId.
        /// </summary>
        /// <param name="combinedSleeve">Combined sleeve data to save</param>
        /// <returns>Database-generated CombinedSleeveId</returns>
        int SaveCombinedSleeve(CombinedSleeve combinedSleeve);
        
        /// <summary>
        /// Saves constituents for a combined sleeve.
        /// Used when constituents are added after initial combined sleeve creation.
        /// </summary>
        /// <param name="combinedSleeveId">ID of the combined sleeve</param>
        /// <param name="constituents">List of constituents to save</param>
        void SaveConstituents(int combinedSleeveId, List<SleeveConstituent> constituents);
        
        // ============================================================================
        // READ OPERATIONS
        // ============================================================================
        
        /// <summary>
        /// Retrieves a combined sleeve by its database ID.
        /// </summary>
        /// <param name="combinedSleeveId">Database ID</param>
        /// <returns>Combined sleeve data, or null if not found</returns>
        CombinedSleeve GetCombinedSleeveById(int combinedSleeveId);
        
        /// <summary>
        /// Retrieves a combined sleeve by its Revit instance ID.
        /// </summary>
        /// <param name="instanceId">Revit ElementId.IntegerValue</param>
        /// <returns>Combined sleeve data, or null if not found</returns>
        CombinedSleeve GetCombinedSleeveByInstanceId(int instanceId);
        
        /// <summary>
        /// Retrieves all combined sleeves for a specific combo and filter.
        /// </summary>
        /// <param name="comboId">File combination ID</param>
        /// <param name="filterId">Filter ID</param>
        /// <returns>List of combined sleeves</returns>
        List<CombinedSleeve> GetCombinedSleevesForCombo(int comboId, int filterId);
        
        /// <summary>
        /// Retrieves all constituents for a specific combined sleeve.
        /// </summary>
        /// <param name="combinedSleeveId">Combined sleeve ID</param>
        /// <returns>List of constituents</returns>
        List<SleeveConstituent> GetConstituents(int combinedSleeveId);
        
        // ============================================================================
        // UPDATE OPERATIONS
        // ============================================================================
        
        /// <summary>
        /// Updates the corner coordinates for a combined sleeve.
        /// Follows the same pattern as UpdateSleeveCorners and UpdateClusterSleeveCorners.
        /// </summary>
        /// <param name="combinedSleeveId">Combined sleeve ID</param>
        /// <param name="c1x">Corner 1 X coordinate</param>
        /// <param name="c1y">Corner 1 Y coordinate</param>
        /// <param name="c1z">Corner 1 Z coordinate</param>
        /// <param name="c2x">Corner 2 X coordinate</param>
        /// <param name="c2y">Corner 2 Y coordinate</param>
        /// <param name="c2z">Corner 2 Z coordinate</param>
        /// <param name="c3x">Corner 3 X coordinate</param>
        /// <param name="c3y">Corner 3 Y coordinate</param>
        /// <param name="c3z">Corner 3 Z coordinate</param>
        /// <param name="c4x">Corner 4 X coordinate</param>
        /// <param name="c4y">Corner 4 Y coordinate</param>
        /// <param name="c4z">Corner 4 Z coordinate</param>
        void UpdateCombinedSleeveCorners(int combinedSleeveId,
            double c1x, double c1y, double c1z,
            double c2x, double c2y, double c2z,
            double c3x, double c3y, double c3z,
            double c4x, double c4y, double c4z);
        
        // ============================================================================
        // DELETE OPERATIONS
        // ============================================================================
        
        /// <summary>
        /// Deletes a combined sleeve and all its constituents (cascade delete).
        /// </summary>
        /// <param name="combinedSleeveId">Combined sleeve ID to delete</param>
        void DeleteCombinedSleeve(int combinedSleeveId);
        
        // ============================================================================
        // FLAG OPERATIONS
        // ============================================================================
        
        /// <summary>
        /// Marks constituent sleeves as resolved (sets IsCombinedResolved = true).
        /// Updates both ClashZones and ClusterSleeves tables as appropriate.
        /// </summary>
        /// <param name="constituents">List of constituents to mark as resolved</param>
        void MarkConstituentsAsResolved(List<SleeveConstituent> constituents);
    }
}
