using System.Collections.Generic;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Placement
{
    /// <summary>
    /// Abstraction for pre-placement planning. Implementations must be pure (no Revit API calls).
    /// </summary>
    public interface ISleevePlacementPlanner
    {
        SleevePlacementPlanningResult Plan(IEnumerable<ClashZone> clashZones);
        SleevePlacementPlanningDto PlanSingle(ClashZone zone);
    }
}
