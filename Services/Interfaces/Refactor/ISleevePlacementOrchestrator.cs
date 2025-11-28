using System.Collections.Generic;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Services.Placement;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor
{
    /// <summary>
    /// Orchestrates the sleeve placement pipeline through sequential stage execution.
    /// Implements SOLID principles: depends on abstractions (IPlacementStage), orchestrates workflow without business logic.
    /// </summary>
    public interface ISleevePlacementOrchestrator
    {
        /// <summary>
        /// Execute the placement pipeline for the given clash zones.
        /// </summary>
        /// <param name="doc">Revit document for element operations</param>
        /// <param name="zones">Clash zones requiring sleeve placement</param>
        /// <param name="perf">Performance monitor for timing stages</param>
        /// <returns>Aggregated result with success status and counts</returns>
        OrchestratorResult Execute(Document doc, IReadOnlyList<ClashZone> zones, IPerformanceMonitor perf);
    }
}
