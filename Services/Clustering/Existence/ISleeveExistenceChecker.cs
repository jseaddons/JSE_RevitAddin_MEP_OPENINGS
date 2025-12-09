using System.Collections.Generic;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Existence
{
    /// <summary>
    /// ✅ SOLID SRP: Interface for checking sleeve existence in Revit.
    /// Separates existence checking logic from placement logic.
    /// 
    /// Part of 28 Features from Comprehensive Architecture:
    /// - 🏗️ SOLID Architecture (SRP, DIP)
    /// - 📊 Diagnostic Logging
    /// - 🛡️ Crash-Safe Execution (Exception handling)
    /// </summary>
    public interface ISleeveExistenceChecker
    {
        /// <summary>
        /// Check if a sleeve exists in Revit by SleeveInstanceId.
        /// </summary>
        /// <param name="sleeveInstanceId">Sleeve instance ID from database</param>
        /// <returns>True if sleeve exists and is valid, false otherwise</returns>
        bool SleeveExists(int sleeveInstanceId);

        /// <summary>
        /// Check if a cluster sleeve exists in Revit by ClusterInstanceId.
        /// </summary>
        /// <param name="clusterInstanceId">Cluster instance ID from database</param>
        /// <returns>True if cluster sleeve exists and is valid, false otherwise</returns>
        bool ClusterSleeveExists(int clusterInstanceId);

        /// <summary>
        /// Check if individual sleeves exist for given clash zones.
        /// </summary>
        /// <param name="clashZoneIds">List of clash zone GUIDs</param>
        /// <returns>Dictionary mapping clash zone ID to whether individual sleeve exists</returns>
        Dictionary<System.Guid, bool> CheckIndividualSleevesExist(List<System.Guid> clashZoneIds);

        /// <summary>
        /// Check if any individual sleeves exist that would conflict with a cluster.
        /// </summary>
        /// <param name="clashZoneIds">List of clash zone GUIDs in the cluster</param>
        /// <returns>True if any conflicting individual sleeves exist, false otherwise</returns>
        bool HasConflictingIndividualSleeves(List<System.Guid> clashZoneIds);
    }
}

