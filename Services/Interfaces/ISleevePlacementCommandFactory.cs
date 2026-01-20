using System.Collections.Generic;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Commands;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces
{
    /// <summary>
    /// Interface for creating sleeve placement commands.
    /// Extracted from OpeningCommandOrchestrator to adhere to SRP and DIP.
    /// </summary>
    public interface ISleevePlacementCommandFactory
    {
        /// <summary>
        /// Create a sleeve placement command for the given parameters.
        /// </summary>
        ICommand CreateSleevePlacementCommand(
            Document doc,
            List<ClashZone> clashZones,
            string category,
            string filterName,
            Dictionary<string, double>? clearanceSettings);
    }
}

