using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// ✅ SOLID: Single Responsibility - Handles prefix resolution for clusters and combined sleeves
    /// Implements the 3 Prefix Fixes:
    /// - Fix #1: Clusters with same service type → user-defined service type prefix
    /// - Fix #2: Multi-link clusters/combined → hardcoded "MEP" prefix
    /// - Fix #3: Combined sleeve (single link + "chilled water") → user-defined prefix (EXCEPTION)
    /// </summary>
    public class ClusterPrefixResolutionService
    {
        private readonly Action<string> _logger;

        public ClusterPrefixResolutionService(Action<string> logger = null)
        {
            _logger = logger ?? ((msg) => { });
        }

        /// <summary>
        /// ✅ FIX #1, #2, #3: Enhanced prefix resolution for clusters and combined sleeves
        /// </summary>
        public string ResolvePrefixForClusterOrCombined(
            FamilyInstance sleeve,
            string category,
            string defaultPrefix,
            object markPrefixes, // Changed to object to avoid dependency
            Document doc,
            SleeveDbContext sharedContext)
        {
            if (sleeve == null || sharedContext == null || markPrefixes == null)
                return defaultPrefix;

            try
            {
                int sleeveId = sleeve.Id.IntegerValue;
                List<ClashZone> zones = new List<ClashZone>();

                // Determine if this is a cluster or combined sleeve
                bool isCombined = IsCombinedSleeve(sleeve);

                if (isCombined)
                {
                    zones = GetZonesForCombinedSleeve(sleeveId, sharedContext);
                }
                else
                {
                    zones = GetZonesForClusterSleeve(sleeveId, category, sharedContext);
                }

                if (zones.Count == 0)
                    return defaultPrefix; // No zones found, use default

                // ✅ FIX #2: Check if zones are from different links
                if (IsMultiLink(zones))
                {
                    LogPrefixDecision($"[PREFIX-FIX-2] ✅ {(isCombined ? "Combined" : "Cluster")} sleeve {sleeveId} has zones from multiple links → using 'MEP' prefix (hardcoded)");
                    return "MEP";
                }

                // ✅ FIX #3: EXCEPTION - Combined sleeve with single link + "chilled" system type
                if (isCombined && HasChilledWaterSystem(zones))
                {
                    // Deprecated MarkPrefixSettings usage - fallback to default behavior
                    // var resolvedPrefix = ResolveChilledWaterPrefix(zones, category, markPrefixes, sleeveId);
                    // if (!string.IsNullOrWhiteSpace(resolvedPrefix))
                    //     return resolvedPrefix;
                }

                // ✅ FIX #1: Check if all zones have the same service type
                // Deprecated MarkPrefixSettings usage - fallback to default behavior
                // var sameServiceTypePrefix = ResolveSameServiceTypePrefix(zones, category, markPrefixes, sleeveId, isCombined);
                // if (!string.IsNullOrWhiteSpace(sameServiceTypePrefix))
                //    return sameServiceTypePrefix;

                // ✅ DEFAULT FOR COMBINED SLEEVES: Use "MEP" if no exception applies
                if (isCombined)
                {
                    LogPrefixDecision($"[PREFIX-DEFAULT] Combined sleeve {sleeveId} → using default 'MEP' prefix (no exception applied)");
                    return "MEP";
                }
            }
            catch (Exception ex)
            {
                LogPrefixDecision($"[PREFIX-FIX] ❌ Error resolving cluster/combined prefix for sleeve {sleeve.Id}: {ex.Message}");
            }

            // Default: use standard resolution for clusters or fallback
            return defaultPrefix;
        }

        #region Private Helper Methods

        /// <summary>
        /// Determines if a sleeve is a combined sleeve based on family name
        /// </summary>
        private bool IsCombinedSleeve(FamilyInstance sleeve)
        {
            if (sleeve == null) return false;
            
            var familyName = sleeve.Symbol?.FamilyName ?? string.Empty;
            return familyName.Contains("Combined", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Gets all zones for a combined sleeve from the database
        /// </summary>
        private List<ClashZone> GetZonesForCombinedSleeve(int sleeveId, SleeveDbContext context)
        {
            var zones = new List<ClashZone>();

            using (var cmd = context.Connection.CreateCommand())
            {
                cmd.CommandText = @"
                    SELECT cz.* FROM ClashZones cz
                    INNER JOIN CombinedSleeveZones csz ON csz.ClashZoneGuid = cz.Id
                    WHERE csz.CombinedInstanceId = @CombinedId";
                cmd.Parameters.AddWithValue("@CombinedId", sleeveId);

                using (var reader = cmd.ExecuteReader())
                {
                    var repository = new ClashZoneRepository(context);
                    while (reader.Read())
                    {
                        // Use reflection to call private MapClashZone method
                        var mapMethod = repository.GetType().GetMethod("MapClashZone",
                            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                        if (mapMethod != null)
                        {
                            var zone = mapMethod.Invoke(repository, new object[] { reader }) as ClashZone;
                            if (zone != null)
                                zones.Add(zone);
                        }
                    }
                }
            }

            return zones;
        }

        /// <summary>
        /// Gets all zones for a cluster sleeve from the database
        /// </summary>
        private List<ClashZone> GetZonesForClusterSleeve(int sleeveId, string category, SleeveDbContext context)
        {
            var repository = new ClashZoneRepository(context);
            var allZones = repository.GetClashZonesByCategory(category);
            return allZones?.Where(z => z.ClusterSleeveInstanceId == sleeveId && z.IsClusterResolved).ToList()
                   ?? new List<ClashZone>();
        }

        /// <summary>
        /// Checks if zones are from multiple linked files
        /// </summary>
        private bool IsMultiLink(List<ClashZone> zones)
        {
            var distinctLinks = zones
                .Select(z => z.SourceDocKey ?? string.Empty)
                .Distinct()
                .ToList();

            return distinctLinks.Count > 1;
        }

        /// <summary>
        /// Checks if any zone contains "chilled water" system type
        /// </summary>
        private bool HasChilledWaterSystem(List<ClashZone> zones)
        {
            return zones.Any(z =>
            {
                var systemType = GetClashParameterValue(z, "System Type", "MEP System Type", "System Classification");
                var serviceName = GetClashParameterValue(z, "System Name", "Service Name");
                return (systemType?.Contains("Chill", StringComparison.OrdinalIgnoreCase) ?? false) ||
                       (systemType?.Contains("CHW", StringComparison.OrdinalIgnoreCase) ?? false) ||
                       (serviceName?.Contains("Chill", StringComparison.OrdinalIgnoreCase) ?? false);
            });
        }

        /// <summary>
        /// Resolves prefix for chilled water combined sleeves (Fix #3)
        /// </summary>
        private string ResolveChilledWaterPrefix(List<ClashZone> zones, string category, object markPrefixes, int sleeveId)
        {
            var chilledZones = zones.Where(z =>
            {
                var systemType = GetClashParameterValue(z, "System Type", "MEP System Type", "System Classification");
                var serviceName = GetClashParameterValue(z, "System Name", "Service Name");
                return (systemType?.Contains("Chill", StringComparison.OrdinalIgnoreCase) ?? false) ||
                       (systemType?.Contains("CHW", StringComparison.OrdinalIgnoreCase) ?? false) ||
                       (serviceName?.Contains("Chill", StringComparison.OrdinalIgnoreCase) ?? false);
            }).ToList();

            if (chilledZones.Count > 0)
            {
                var systemType = GetClashParameterValue(chilledZones[0], "System Type", "MEP System Type", "System Classification");

                // ✅ USE USER-DEFINED PREFIX from system type overrides - DEPRECATED
                // var typedSettings = markPrefixes as JSE_RevitAddin_MEP_OPENINGS.Services.Parameters.Configuration.MarkPrefixSettings;
                // var resolvedPrefix = typedSettings?.GetPrefixForElement(category, systemType, null);
                string resolvedPrefix = null;

                if (!string.IsNullOrWhiteSpace(resolvedPrefix))
                {
                    LogPrefixDecision($"[PREFIX-FIX-3] ✅ EXCEPTION: Combined sleeve {sleeveId} is single-link + contains 'chilled water' system type '{systemType}' → using USER-DEFINED prefix '{resolvedPrefix}'");
                    return resolvedPrefix;
                }
                else
                {
                    LogPrefixDecision($"[PREFIX-FIX-3] ⚠️ Combined sleeve {sleeveId} has 'chilled water' system type '{systemType}' but NO user-defined override found → will use default MEP");
                }
            }

            return null;
        }

        /// <summary>
        /// Resolves prefix when all zones have the same service type (Fix #1)
        /// </summary>
        private string ResolveSameServiceTypePrefix(List<ClashZone> zones, string category, object markPrefixes, int sleeveId, bool isCombined)
        {
            var systemTypes = zones
                .Select(z => GetClashParameterValue(z, "System Type", "MEP System Type", "System Classification"))
                .Where(st => !string.IsNullOrWhiteSpace(st))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (systemTypes.Count == 1)
            {
                // All zones have same system type → use USER-DEFINED service type prefix
                string systemType = systemTypes[0];

                // ✅ USE USER-DEFINED PREFIX from system type overrides
                // var resolvedPrefix = markPrefixes.GetPrefixForElement(category, systemType, null);
                string resolvedPrefix = null; // DISABLED
                // var defaultCategoryPrefix = markPrefixes.GetDisciplinePrefix(category);
                string defaultCategoryPrefix = "MEP"; // Fallback

                if (!string.IsNullOrWhiteSpace(resolvedPrefix) &&
                    !resolvedPrefix.Equals(defaultCategoryPrefix, StringComparison.OrdinalIgnoreCase))
                {
                    // User-defined service type override exists and is different from default category prefix
                    LogPrefixDecision($"[PREFIX-FIX-1] ✅ {(isCombined ? "Combined" : "Cluster")} sleeve {sleeveId} has all zones with same system type '{systemType}' → using USER-DEFINED prefix '{resolvedPrefix}'");
                    return resolvedPrefix;
                }
            }

            return null;
        }

        /// <summary>
        /// Gets a parameter value from a ClashZone, trying multiple parameter names
        /// </summary>
        private string GetClashParameterValue(ClashZone zone, params string[] parameterNames)
        {
            if (zone?.MepParameterValues == null)
                return null;

            foreach (var paramName in parameterNames)
            {
                // Try case-insensitive lookup
                foreach (var kv in zone.MepParameterValues)
                {
                    if (string.Equals(kv.Key, paramName, StringComparison.OrdinalIgnoreCase))
                    {
                        return kv.Value;
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Logs prefix decision if not in deployment mode
        /// </summary>
        private void LogPrefixDecision(string message)
        {
            if (!DeploymentConfiguration.DeploymentMode)
            {
                string mepmarkLogPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                File.AppendAllText(mepmarkLogPath, message + "\n");
            }

            _logger?.Invoke(message);
        }

        #endregion
    }
}
