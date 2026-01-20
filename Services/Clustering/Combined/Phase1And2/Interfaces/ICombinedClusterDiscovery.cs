using System;
using System.Collections.Generic;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Interfaces
{
    /// <summary>
    /// Legacy compatibility shim for early combined clustering discovery contracts.
    /// Prefer ICombinedClusterDiscoveryService.
    /// </summary>
    [Obsolete("Use ICombinedClusterDiscoveryService instead.")]
    public interface ICombinedClusterDiscovery
    {
        IReadOnlyList<ClusterSleeveInfo> Discover(string filterName, IReadOnlyCollection<string> categories);
    }
}
