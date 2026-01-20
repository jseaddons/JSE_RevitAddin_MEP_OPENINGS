using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor;
using JSE_RevitAddin_MEP_OPENINGS.Utils;
using JSE_RevitAddin_MEP_OPENINGS.Data;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Placement.Stages
{
    /// <summary>
    /// Stage 1: Filter clash zones based on ReadyForPlacement flags and section box.
    /// Uses Team C's IZonePreFilterService with section box filtering when available.
    /// </summary>
    public class FilterStage : IPlacementStage
    {
        public string Name => "FilterZones";
        private readonly IZonePreFilterService _preFilterService;

        public FilterStage(IZonePreFilterService preFilterService = null)
        {
            _preFilterService = preFilterService; // optional; falls back to simple predicate
        }

        public StageResult Execute(PlacementContext context, IPerformanceMonitor perf)
        {
            try
            {
                perf?.StartOperation(Name);

                var input = context.InputZones;

                // Build criteria with section box (if active 3D view)
                BoundingBoxXYZ sectionBox = null;
                if (context.Doc?.ActiveView is Autodesk.Revit.DB.View3D v3 && v3.IsSectionBoxActive)
                {
                    sectionBox = JSE_RevitAddin_MEP_OPENINGS.Helpers.SectionBoxHelper.GetSectionBoxBounds(v3);
                }

                var criteria = new FilterCriteria
                {
                    UseRTreeIndex = true,
                    UseSpatialGrid = true,
                    SectionBox = sectionBox,
                    CustomPredicate = z => z.ReadyForPlacement,
                    ExcludePlacedZones = true
                };

                List<ClashZone> filtered;
                if (_preFilterService != null)
                {
                    // ✅ FIX: Convert IReadOnlyList to IList for Filter method
                    var inputList = input.ToList();
                    var filteredResult = _preFilterService.Filter(inputList, criteria);
                    // ✅ FIX: Ensure we have a List<T> (not just IList<T>) for type compatibility
                    filtered = filteredResult as List<ClashZone> ?? filteredResult.ToList();
                }
                else
                {
                    // Fallback: simple predicate + optional bounds check
                    filtered = input
                        .Where(z => z.ReadyForPlacement)
                        .Where(z => sectionBox == null ||
                                    (z.IntersectionPoint != null &&
                                     z.IntersectionPoint.X >= sectionBox.Min.X && z.IntersectionPoint.X <= sectionBox.Max.X &&
                                     z.IntersectionPoint.Y >= sectionBox.Min.Y && z.IntersectionPoint.Y <= sectionBox.Max.Y &&
                                     z.IntersectionPoint.Z >= sectionBox.Min.Z && z.IntersectionPoint.Z <= sectionBox.Max.Z))
                        .ToList();
                }

                perf?.StopOperation(Name, filtered.Count);

                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[{Name}] Filtered {context.InputZones.Count} -> {filtered.Count} zones");
                }

                // ✅ FIX: List<T> implements IReadOnlyList<T>, so we can pass it directly
                var updatedContext = context
                    .WithFilteredZones(filtered)
                    .WithMetric("FilteredCount", filtered.Count)
                    .WithMetric("FilteredOutCount", context.InputZones.Count - filtered.Count);

                return StageResult.Succeeded(updatedContext);
            }
            catch (Exception ex)
            {
                return StageResult.Failed(context, $"FilterStage failed: {ex.Message}");
            }
        }
    }
}
