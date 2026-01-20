using System;
using System.Collections.Generic;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Rotation.Interfaces;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Rotation.Services
{
    /// <summary>
    /// ✅ SOLID SRP: Service responsible ONLY for caching cluster rotation data
    /// Single Responsibility: Manage cache for ClashZone lookups and rotated bounding box calculations
    /// </summary>
    public class ClusterRotationCacheService : IClusterRotationCacheService
    {
        // ✅ PERFORMANCE: Cache individual ClashZone lookups to avoid repeated database queries
        private readonly Dictionary<int, ClashZone> _clashZoneCache;
        private const int MAX_CLASHZONE_CACHE_SIZE = 5000;
        
        // Delegate for getting ClashZone by sleeve instance ID (injected dependency)
        private readonly Func<int, string, ClashZone> _getClashZoneFunc;

        public ClusterRotationCacheService(Func<int, string, ClashZone> getClashZoneFunc)
        {
            if (getClashZoneFunc == null)
                throw new ArgumentNullException(nameof(getClashZoneFunc));
            
            _clashZoneCache = new Dictionary<int, ClashZone>();
            _getClashZoneFunc = getClashZoneFunc;
        }

        /// <summary>
        /// ✅ PERFORMANCE: Get ClashZone with caching to avoid repeated database queries
        /// </summary>
        public ClashZone? GetCachedClashZone(int sleeveInstanceId, string? xmlFilePath)
        {
            if (sleeveInstanceId <= 0)
                return null;
            
            // Check cache first
            if (_clashZoneCache.TryGetValue(sleeveInstanceId, out var cached))
            {
                return cached;
            }
            
            // Cache miss - load from function
            var cz = _getClashZoneFunc(sleeveInstanceId, xmlFilePath ?? string.Empty);
            var clashZone = cz as ClashZone;
            
            // Store in cache (if not null and cache not full)
            if (clashZone != null && _clashZoneCache.Count < MAX_CLASHZONE_CACHE_SIZE)
            {
                _clashZoneCache[sleeveInstanceId] = clashZone;
            }
            
            return clashZone;
        }

        /// <summary>
        /// Clear all caches
        /// </summary>
        public void ClearCaches()
        {
            _clashZoneCache.Clear();
        }
    }
}

