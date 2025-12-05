using System.Collections.Generic;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Rotation.Interfaces
{
    /// <summary>
    /// ✅ SOLID SRP: Service responsible ONLY for determining cluster rotation angles
    /// Single Responsibility: Calculate rotation angle based on cluster composition
    /// </summary>
    public interface IClusterRotationAngleService
    {
        /// <summary>
        /// Determine the dominant rotation angle for a cluster of sleeves
        /// </summary>
        /// <param name="cluster">List of sleeve data (dynamic objects with ClashZone property)</param>
        /// <param name="xmlFilePath">Optional XML file path for data access</param>
        /// <returns>Rotation angle in radians (0 for axis-aligned clusters)</returns>
        double DetermineRotationAngle(List<dynamic> cluster, string? xmlFilePath = null);
    }
}

