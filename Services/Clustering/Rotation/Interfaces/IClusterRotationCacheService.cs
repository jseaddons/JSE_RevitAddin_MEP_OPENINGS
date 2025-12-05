using System.Collections.Generic;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Rotation.Interfaces
{
    /// <summary>
    /// ✅ SOLID SRP: Service responsible ONLY for caching cluster rotation data
    /// Single Responsibility: Manage cache for ClashZone lookups and rotated bounding box calculations
    /// </summary>
    public interface IClusterRotationCacheService
    {
        /// <summary>
        /// Get cached ClashZone by sleeve instance ID
        /// </summary>
        ClashZone? GetCachedClashZone(int sleeveInstanceId, string? xmlFilePath);
        
        /// <summary>
        /// Clear all caches
        /// </summary>
        void ClearCaches();
    }
}

