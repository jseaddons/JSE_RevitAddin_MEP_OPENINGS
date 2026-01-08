using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Models;

using JSE_RevitAddin_MEP_OPENINGS.Services.Parameters.Configuration;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Parameters.Processing
{
    public interface IMarkCacheService
    {
        void Initialize(Document doc, SleeveDbContext sharedContext);
        ClashZone GetClashZoneByMepElementId(long mepElementId);
        ClashZone GetClashZoneByClusterInstanceId(int clusterInstanceId);
        bool IsCombinedSleeve(int instanceId);
    }

    public class MarkCacheService : IMarkCacheService
    {
        private Dictionary<long, ClashZone> _clashZoneCache = new Dictionary<long, ClashZone>();
        private Dictionary<int, ClashZone> _clusterZoneCache = new Dictionary<int, ClashZone>();
        private HashSet<int> _combinedSleeveCache = new HashSet<int>();

        public void Initialize(Document doc, SleeveDbContext sharedContext)
        {
            _clashZoneCache.Clear();
            _clusterZoneCache.Clear();
            _combinedSleeveCache.Clear();

            try
            {
                if (sharedContext != null)
                {
                    var repository = new ClashZoneRepository(sharedContext);
                    var categories = new[] { "Ducts", "Pipes", "Cable Trays", "Duct Accessories" };

                    foreach (var category in categories)
                    {
                        var zones = repository.GetClashZonesByCategory(category);
                        if (zones == null) continue;

                        foreach (var cz in zones)
                        {
                            if (cz == null) continue;
                            if (cz.MepElementIdValue > 0 && !_clashZoneCache.ContainsKey(cz.MepElementIdValue))
                            {
                                cz.EnsureSleevePlacementPointReconstructed();
                                _clashZoneCache[cz.MepElementIdValue] = cz;
                            }
                            if (cz.IsClusterResolved && cz.ClusterSleeveInstanceId > 0)
                            {
                                _clusterZoneCache[cz.ClusterSleeveInstanceId] = cz;
                            }
                        }
                    }

                    // Combined Sleeves Cache
                    try
                    {
                        var comboRepo = new CombinedSleeveRepository(sharedContext);
                        foreach (var combo in comboRepo.GetAllCombinedSleeves())
                        {
                            if (combo.CombinedInstanceId > 0)
                                _combinedSleeveCache.Add(combo.CombinedInstanceId);
                        }
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[MarkCacheService] Database initialization failed: {ex.Message}");
            }
        }

        public ClashZone GetClashZoneByMepElementId(long mepElementId) => 
            _clashZoneCache.TryGetValue(mepElementId, out var cz) ? cz : null;

        public ClashZone GetClashZoneByClusterInstanceId(int clusterInstanceId) => 
            _clusterZoneCache.TryGetValue(clusterInstanceId, out var cz) ? cz : null;

        public bool IsCombinedSleeve(int instanceId) => _combinedSleeveCache.Contains(instanceId);
    }
}
