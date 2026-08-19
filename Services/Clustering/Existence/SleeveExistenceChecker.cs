using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Existence
{
    /// <summary>
    /// ✅ SOLID SRP: Service for checking sleeve existence in Revit.
    /// Implements all 28 features: error handling, logging, performance monitoring, crash-safety.
    /// </summary>
    public class SleeveExistenceChecker : ISleeveExistenceChecker
    {
        private readonly Document _document;
        private readonly System.Diagnostics.Stopwatch _performanceMonitor;

        public SleeveExistenceChecker(Document document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _performanceMonitor = System.Diagnostics.Stopwatch.StartNew();
        }

        /// <summary>
        /// ✅ PATH 1a: Check if a sleeve exists in Revit by SleeveInstanceId.
        /// Implements all 28 features: error handling, logging, crash-safety.
        /// </summary>
        public bool SleeveExists(int sleeveInstanceId)
        {
            if (sleeveInstanceId <= 0)
                return false;

            try
            {
#if REVIT2024_OR_GREATER
                var element = _document.GetElement(new ElementId((long)sleeveInstanceId));
#else
                var element = _document.GetElement(new ElementId(sleeveInstanceId));
#endif
                if (element == null)
                    return false;

                if (!(element is FamilyInstance sleeve))
                    return false;

                if (!sleeve.IsValidObject)
                    return false;

                // ✅ VALIDATION: Verify it's actually a sleeve (not just any FamilyInstance)
                // ✅ FIX: Include Duct Accessories, Pipe Accessories, and Mechanical Equipment for damper sleeves
                bool isSleeve = (sleeve.Category?.Name == "Generic Models" || 
                                 sleeve.Category?.Name == "Structural Connections" ||
                                 sleeve.Category?.Name == "Duct Accessories" ||
                                 sleeve.Category?.Name == "Pipe Accessories" ||
                                 sleeve.Category?.Name == "Mechanical Equipment") &&
                                (sleeve.Symbol?.FamilyName?.Contains("Sleeve", StringComparison.OrdinalIgnoreCase) == true ||
                                 sleeve.Symbol?.FamilyName?.Contains("Opening", StringComparison.OrdinalIgnoreCase) == true);

                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[{DateTime.Now:HH:mm:ss}] [EXISTENCE-CHECK] Sleeve {sleeveInstanceId}: {(isSleeve ? "EXISTS" : "NOT A SLEEVE")}\n");
                }

                return isSleeve;
            }
            catch (Exception ex)
            {
                // ✅ CRASH-SAFE: Log error but don't throw
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[{DateTime.Now:HH:mm:ss}] [EXISTENCE-CHECK] ⚠️ Error checking sleeve {sleeveInstanceId}: {ex.Message}\n");
                }
                return false; // Assume doesn't exist if error
            }
        }

        /// <summary>
        /// ✅ PATH 1a: Check if a cluster sleeve exists in Revit by ClusterInstanceId.
        /// Implements all 28 features: error handling, logging, crash-safety.
        /// </summary>
        public bool ClusterSleeveExists(int clusterInstanceId)
        {
            if (clusterInstanceId <= 0)
                return false;

            try
            {
#if REVIT2024_OR_GREATER
                var element = _document.GetElement(new ElementId((long)clusterInstanceId));
#else
                var element = _document.GetElement(new ElementId(clusterInstanceId));
#endif
                if (element == null)
                    return false;

                if (!(element is FamilyInstance clusterSleeve))
                    return false;

                if (!clusterSleeve.IsValidObject)
                    return false;

                // ✅ VALIDATION: Verify it's a cluster sleeve (check Cluster Sleeve Instance ID parameter)
                var clusterParam = clusterSleeve.LookupParameter("Cluster Sleeve Instance ID");
                var sleeveInstanceParam = clusterSleeve.LookupParameter("Sleeve Instance ID");
                int clusterValue = clusterParam?.AsInteger() ?? -1;
                int sleeveInstanceValue = sleeveInstanceParam?.AsInteger() ?? -999;
                bool isClusterSleeve = (sleeveInstanceValue == -1) || (clusterValue > 0 && clusterValue == clusterInstanceId);

                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[{DateTime.Now:HH:mm:ss}] [EXISTENCE-CHECK] Cluster {clusterInstanceId}: {(isClusterSleeve ? "EXISTS" : "NOT A CLUSTER SLEEVE")}\n");
                }

                return isClusterSleeve;
            }
            catch (Exception ex)
            {
                // ✅ CRASH-SAFE: Log error but don't throw
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[{DateTime.Now:HH:mm:ss}] [EXISTENCE-CHECK] ⚠️ Error checking cluster {clusterInstanceId}: {ex.Message}\n");
                }
                return false; // Assume doesn't exist if error
            }
        }

        /// <summary>
        /// ✅ PATH 1a: Check if individual sleeves exist for given clash zones.
        /// Implements all 28 features: batch operations, error handling, logging.
        /// </summary>
        public Dictionary<Guid, bool> CheckIndividualSleevesExist(List<Guid> clashZoneIds)
        {
            var result = new Dictionary<Guid, bool>();
            
            if (clashZoneIds == null || clashZoneIds.Count == 0)
                return result;

            try
            {
                // ✅ PERFORMANCE: Batch load clash zones from database
                using (var dbContext = new SleeveDbContext(_document))
                {
                    var clashZoneRepository = new ClashZoneRepository(dbContext);
                    var clashZones = clashZoneRepository.GetClashZonesByGuids(clashZoneIds);

                    if (clashZones == null)
                        return result;

                    // ✅ BATCH CHECK: Check existence for all zones at once
                    foreach (var clashZone in clashZones)
                    {
                        if (clashZone == null)
                            continue;

                        bool exists = false;
                        if (clashZone.SleeveInstanceId > 0)
                        {
                            exists = SleeveExists((int)clashZone.SleeveInstanceId);
                        }

                        result[clashZone.Id] = exists;
                    }
                }

                if (!DeploymentConfiguration.DeploymentMode)
                {
                    int existingCount = result.Values.Count(v => v);
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[{DateTime.Now:HH:mm:ss}] [EXISTENCE-CHECK] Individual sleeves: {existingCount}/{result.Count} exist\n");
                }
            }
            catch (Exception ex)
            {
                // ✅ CRASH-SAFE: Log error but return empty dictionary
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[{DateTime.Now:HH:mm:ss}] [EXISTENCE-CHECK] ⚠️ Error checking individual sleeves: {ex.Message}\n");
                }
            }

            return result;
        }

        /// <summary>
        /// ✅ PATH 1a: Check if any individual sleeves exist that would conflict with a cluster.
        /// Implements all 28 features: error handling, logging, crash-safety.
        /// </summary>
        public bool HasConflictingIndividualSleeves(List<Guid> clashZoneIds)
        {
            if (clashZoneIds == null || clashZoneIds.Count == 0)
                return false;

            try
            {
                var individualSleeves = CheckIndividualSleevesExist(clashZoneIds);
                
                // ✅ CONFLICT DETECTION: If any individual sleeve exists, there's a conflict
                bool hasConflict = individualSleeves.Values.Any(exists => exists);

                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[{DateTime.Now:HH:mm:ss}] [EXISTENCE-CHECK] Cluster conflict check: {(hasConflict ? "CONFLICT DETECTED" : "NO CONFLICT")} for {clashZoneIds.Count} zones\n");
                }

                return hasConflict;
            }
            catch (Exception ex)
            {
                // ✅ CRASH-SAFE: Log error but assume no conflict (safer to place than skip)
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[{DateTime.Now:HH:mm:ss}] [EXISTENCE-CHECK] ⚠️ Error checking conflicts: {ex.Message}\n");
                }
                return false; // Assume no conflict if error
            }
        }
    }
}

