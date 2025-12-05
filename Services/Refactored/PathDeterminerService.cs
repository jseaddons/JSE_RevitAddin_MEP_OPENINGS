using System.Collections.Generic;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Refactored
{
    /// <summary>
    /// ✅ SOLID REFACTORED: Service for determining sleeve placement path (PATH 1, PATH 2, PATH 3).
    /// Extracted from UniversalSleevePlacementCommand to adhere to SRP.
    /// </summary>
    public class PathDeterminerService : IPathDeterminer
    {
        public SleevePlacementPath DeterminePath(
            Document doc,
            string filterName,
            string category,
            string logPrefix,
            Dictionary<string, double>? clearanceSettings)
        {
            // Convert OpeningConditions.ClearanceSettings to Dictionary<string, double> for comparison
            Dictionary<string, double> currentClearanceSettings = null;
            if (clearanceSettings != null)
            {
                currentClearanceSettings = new Dictionary<string, double>(clearanceSettings);
            }

            return Services.Refresh.RefreshPathDeterminer.DeterminePlacementPath(
                doc,
                filterName,
                category,
                logPrefix,
                currentClearanceSettings);
        }
    }
}

