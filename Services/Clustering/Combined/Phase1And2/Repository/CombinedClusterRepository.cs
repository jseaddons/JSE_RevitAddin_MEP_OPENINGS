using System;
using System.Collections.Generic;
using System.Linq;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Repository
{
    /// <summary>
    /// Repository wrapper that exposes cluster sleeves and parameter snapshots for combined clustering.
    /// </summary>
    public class CombinedClusterRepository : ICombinedClusterRepository
    {
        private readonly IClashZoneRepository _clashZoneRepository;

        public CombinedClusterRepository(IClashZoneRepository clashZoneRepository)
        {
            _clashZoneRepository = clashZoneRepository ?? throw new ArgumentNullException(nameof(clashZoneRepository));
        }

        public IReadOnlyList<ClashZone> LoadClusteredZones(string filterName, IReadOnlyCollection<string> categories)
        {
            if (categories == null || categories.Count == 0)
            {
                return Array.Empty<ClashZone>();
            }

            var resolved = new List<ClashZone>();
            // ✅ AUTO-DISCOVERY FIX: Handle wildcard or missing filter name by scanning all filters
            bool scanAllFilters = string.IsNullOrWhiteSpace(filterName) || filterName.Trim() == "*";

            foreach (var category in categories)
            {
                List<ClashZone> zones;
                if (scanAllFilters)
                {
                    zones = _clashZoneRepository.GetClashZonesByCategory(category);
                }
                else
                {
                    zones = _clashZoneRepository.GetClashZonesByFilter(filterName, category, unresolvedOnly: false, readyForPlacementOnly: false);
                }

                if (zones == null || zones.Count == 0)
                {
                    continue;
                }

                resolved.AddRange(zones.Where(z => z.IsClusterResolved && z.ClusterSleeveInstanceId > 0));
            }

            return resolved;
        }

        public Dictionary<int, Dictionary<string, string>> LoadSnapshotParameters(IEnumerable<int> sleeveInstanceIds)
        {
            var ids = sleeveInstanceIds?.Where(id => id > 0).Distinct().ToList();
            if (ids == null || ids.Count == 0)
            {
                return new Dictionary<int, Dictionary<string, string>>();
            }

            return _clashZoneRepository.GetSnapshotMepParametersForSleeveIds(ids);
        }

        public void SaveCombinedClusterMetadata(string filterName, IReadOnlyCollection<ClashZone> updatedZones)
        {
            if (updatedZones == null || updatedZones.Count == 0)
            {
                return;
            }

            var filter = filterName ?? string.Empty;
            foreach (var group in updatedZones.GroupBy(z => z.MepElementCategory ?? string.Empty))
            {
                _clashZoneRepository.InsertOrUpdateClashZones(group.ToList(), filter, group.Key);
            }
        }

        public List<ClashZone> FindIndividualSleevesNearBoundingBox(Autodesk.Revit.DB.BoundingBoxXYZ boundingBox, string hostType, double level, double toleranceFt)
        {
            // Implementation: Retrieve all candidate zones and filter in memory
            // Candidates: Placed (Resolved) but NOT Clustered (IsClusterResolved = false)
            // We search across all relevant categories.
            
            var categories = new[] { "Ducts", "Pipes", "Cable Trays", "Conduits", "Duct Accessories" };
            var allCandidates = new List<ClashZone>();

            foreach (var cat in categories)
            {
                var zones = _clashZoneRepository.GetClashZonesByCategory(cat);
                if (zones != null)
                {
                    // Filter for placed individual sleeves
                    allCandidates.AddRange(zones.Where(z => z.IsResolved && !z.IsClusterResolved));
                }
            }

            var result = new List<ClashZone>();
            foreach (var zone in allCandidates)
            {
                // Basic overlapping check
                // Zone placement point vs BoundingBox + tolerance
                // ClashZone stores "Active" coordinates if placed
                
                // If zone.SleevePlacementPointActiveX is not set/0, fallback to raw point? 
                // Let's assume SleevePlacementPointX/Y/Z are valid.
                
                // Simple point-in-box (expanded) check
                if (zone.SleevePlacementPointX >= boundingBox.Min.X - toleranceFt &&
                    zone.SleevePlacementPointX <= boundingBox.Max.X + toleranceFt &&
                    zone.SleevePlacementPointY >= boundingBox.Min.Y - toleranceFt &&
                    zone.SleevePlacementPointY <= boundingBox.Max.Y + toleranceFt &&
                    zone.SleevePlacementPointZ >= boundingBox.Min.Z - toleranceFt &&
                    zone.SleevePlacementPointZ <= boundingBox.Max.Z + toleranceFt)
                {
                    result.Add(zone);
                }
            }
            return result;
        }
    }
}
