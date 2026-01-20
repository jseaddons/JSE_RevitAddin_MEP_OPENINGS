using System.Collections.Generic;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces
{
    /// <summary>
    /// Interface for filtering eligible clash zones.
    /// Extracted from UniversalSleevePlacerService to adhere to SRP.
    /// </summary>
    public interface IZoneFilterService
    {
        /// <summary>
        /// Pre-filter eligible clash zones based on UI selection and flags.
        /// </summary>
        List<ClashZone> PreFilterEligibleClashZones(Document doc, List<ClashZone> allClashZones);
    }
}
