using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Repository;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Services
{
    /// <summary>
    /// Phase 1-2 discovery service: gathers cluster sleeves and enriches them with parameter snapshots.
    /// </summary>
    public class CombinedClusterDiscoveryService : ICombinedClusterDiscoveryService
    {
        private readonly ICombinedClusterRepository _repository;

        public CombinedClusterDiscoveryService(ICombinedClusterRepository repository)
        {
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
        }

        public IReadOnlyList<ClusterSleeveInfo> Discover(string filterName, IReadOnlyCollection<string> categories)
        {
            if (!OptimizationFlags.UseCombinedClustering || !OptimizationFlags.UseCombinedClusteringPhase1And2)
            {
                return Array.Empty<ClusterSleeveInfo>();
            }

            var zones = _repository.LoadClusteredZones(filterName, categories);
            if (zones == null || zones.Count == 0)
            {
                return Array.Empty<ClusterSleeveInfo>();
            }

            var sleeves = new ConcurrentBag<ClusterSleeveInfo>();
            System.Threading.Tasks.Parallel.ForEach(zones, BuildParallelOptions(), zone =>
            {
                var sleeve = ClusterSleeveInfo.FromClashZone(zone);
                if (sleeve != null)
                {
                    sleeves.Add(sleeve);
                }
            });

            var materialized = sleeves.ToList();
            if (materialized.Count == 0)
            {
                return Array.Empty<ClusterSleeveInfo>();
            }

            var snapshotLookup = _repository.LoadSnapshotParameters(materialized.Select(s => s.SourceSleeveInstanceId));
            foreach (var sleeve in materialized)
            {
                if (sleeve.SourceSleeveInstanceId > 0 && snapshotLookup.TryGetValue(sleeve.SourceSleeveInstanceId, out var parameters))
                {
                    sleeve.ReplaceParameterSnapshot(parameters);
                }
            }

            return materialized
                .OrderBy(s => s.ClusterSleeveInstanceId)
                .ThenBy(s => s.SourceSleeveInstanceId)
                .ToList();
        }

        private static ParallelOptions BuildParallelOptions()
        {
            var max = OptimizationFlags.CombinedClusterMaxDegreeOfParallelism;
            if (max == 0)
            {
                max = Math.Max(1, Environment.ProcessorCount - 1);
            }

            return new ParallelOptions { MaxDegreeOfParallelism = max };
        }
    }
}
