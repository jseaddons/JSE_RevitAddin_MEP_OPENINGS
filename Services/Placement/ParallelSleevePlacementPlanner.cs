using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Placement;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Placement
{
    /// <summary>
    /// Parallel implementation of sleeve placement planning. Performs pure computations
    /// (sizes, clearance classification, skip rules, rotation heuristics) prior to sequential Revit operations.
    /// </summary>
    public class ParallelSleevePlacementPlanner : ISleevePlacementPlanner
    {
        private readonly int _minParallelCount;
        private readonly int _maxDegree;

        public ParallelSleevePlacementPlanner(int minParallelCount = 12, int? maxDegree = null)
        {
            _minParallelCount = minParallelCount < 1 ? 1 : minParallelCount;
            _maxDegree = maxDegree ?? Environment.ProcessorCount;
        }

        public SleevePlacementPlanningResult Plan(IEnumerable<ClashZone> clashZones)
        {
            var list = clashZones?.ToList() ?? new List<ClashZone>();
            if (list.Count == 0)
            {
                return new SleevePlacementPlanningResult(Array.Empty<SleevePlacementPlanningDto>(), 0, 0, 0, 0);
            }

            var sw = System.Diagnostics.Stopwatch.StartNew();
            var bag = new ConcurrentBag<SleevePlacementPlanningDto>();
            int skipped = 0; int highRisk = 0; int criticalRisk = 0;

            Action<ClashZone> work = zone =>
            {
                try
                {
                    if (zone == null)
                    {
                        System.Threading.Interlocked.Increment(ref skipped);
                        return;
                    }

                    // Basic skip rules (no mutation of original object)
                    string skipReason = null;
                    if (zone.IsResolved && zone.SleeveInstanceId > 0) skipReason = "AlreadyResolved";
                    else if (zone.IsClusterResolved && zone.ClusterSleeveInstanceId > 0) skipReason = "ClusterResolved";
                    else if (zone.HasDamperNearby) skipReason = "DamperSkip";

                    // Host type string
                    string hostType = zone.StructuralElementType ?? string.Empty;

                    // Raw MEP size (convert from internal units to feet)
                    double rawSize = 0.0;
                    if (zone.MepElementSizeData != null)
                    {
                        // MepElementSize stores in internal units, convert to feet
                        if (zone.MepElementSizeData.Diameter > 0)
                        {
                            rawSize = UnitUtils.ConvertFromInternalUnits(zone.MepElementSizeData.Diameter, UnitTypeId.Feet);
                        }
                        else if (zone.MepElementSizeData.Width > 0)
                        {
                            // For rectangular, use width as base size
                            rawSize = UnitUtils.ConvertFromInternalUnits(zone.MepElementSizeData.Width, UnitTypeId.Feet);
                        }
                    }
                    // Fallback to MepElementSize (already in feet)
                    if (rawSize <= 0) rawSize = zone.MepElementSize;
                    if (rawSize <= 0) rawSize = 0.25; // Minimum fallback 3"

                    // Insulation thickness (convert from internal units to feet)
                    double insulation = 0.0;
                    if (zone.MepElementSizeData != null && zone.MepElementSizeData.InsulationThickness > 0)
                    {
                        insulation = UnitUtils.ConvertFromInternalUnits(zone.MepElementSizeData.InsulationThickness, UnitTypeId.Feet);
                    }

                    // Nominal target: add insulation + 0.125ft (1.5") clearance ring
                    double targetWidth = rawSize + insulation + (1.5 / 12.0);
                    double targetHeight = targetWidth; // Assume circular for now; adapt later if rectangular

                    // Host thickness heuristics
                    double hostThickness = zone.StructuralElementThickness;
                    if (hostThickness <= 0 && hostType.Equals("Wall", StringComparison.OrdinalIgnoreCase)) hostThickness = zone.WallThickness;
                    if (hostThickness <= 0 && hostType.Equals("Structural Framing", StringComparison.OrdinalIgnoreCase)) hostThickness = zone.FramingThickness;
                    if (hostThickness <= 0 && hostType.Equals("Floor", StringComparison.OrdinalIgnoreCase)) hostThickness = 0.833; // 10" default
                    if (hostThickness <= 0) hostThickness = 0.5; // Generic fallback 6"

                    // Clearance heuristic: scale with raw size
                    double clearance = Math.Max(0.083, rawSize * 0.10); // >=1" or 10% of size
                    double requiredDepth = hostThickness + clearance;

                    // Rotation angle (convert from radians to degrees)
                    double rotationDeg = 0.0;
                    if (Math.Abs(zone.MepElementRotationAngle) > 1e-6)
                    {
                        rotationDeg = zone.MepElementRotationAngle * 180.0 / Math.PI; // Convert radians to degrees
                    }

                    // Risk classification
                    var risk = ClassifyRisk(hostThickness, clearance, rawSize);
                    if (risk == ClearanceRiskClassification.High) System.Threading.Interlocked.Increment(ref highRisk);
                    if (risk == ClearanceRiskClassification.Critical) System.Threading.Interlocked.Increment(ref criticalRisk);

                    bool shouldSkip = skipReason != null;

                    string logLine = $"PLAN ClashZone={zone.Id} Host={hostType} RawSizeFt={rawSize:F3} TargetFt={targetWidth:F3} ClearanceFt={clearance:F3} DepthFt={requiredDepth:F3} Risk={risk} Skip={shouldSkip} Reason={skipReason}";

                    bag.Add(new SleevePlacementPlanningDto(
                        zone.Id,
                        hostType,
                        rawSize,
                        insulation,
                        targetWidth,
                        targetHeight,
                        clearance,
                        requiredDepth,
                        rotationDeg,
                        risk,
                        shouldSkip,
                        skipReason,
                        logLine));

                    if (shouldSkip) System.Threading.Interlocked.Increment(ref skipped);
                }
                catch
                {
                    System.Threading.Interlocked.Increment(ref skipped);
                }
            };

            if (list.Count >= _minParallelCount)
            {
                Parallel.ForEach(list, new ParallelOptions { MaxDegreeOfParallelism = _maxDegree }, work);
            }
            else
            {
                foreach (var z in list) work(z);
            }

            sw.Stop();
            var ordered = bag.OrderBy(b => b.ShouldSkip).ThenByDescending(b => b.RawMepSizeFt).ToList();
            return new SleevePlacementPlanningResult(ordered, skipped, highRisk, criticalRisk, sw.ElapsedMilliseconds);
        }

        private static ClearanceRiskClassification ClassifyRisk(double hostThickness, double clearance, double rawSize)
        {
            if (rawSize <= 0 || hostThickness <= 0) return ClearanceRiskClassification.Unknown;
            double ratio = clearance / rawSize; // relative clearance
            if (ratio < 0.08) return ClearanceRiskClassification.Critical;  // < ~1" on 1ft size
            if (ratio < 0.10) return ClearanceRiskClassification.High;
            if (ratio < 0.15) return ClearanceRiskClassification.Medium;
            return ClearanceRiskClassification.Low;
        }
    }
}

/*
USAGE (to be wired later inside UniversalSleevePlacerService without modifying existing logic yet):

    // 1. Instantiate once (could be cached)
    var planner = new ParallelSleevePlacementPlanner();

    // 2. Run planning on the clash zones BEFORE sequential placement:
    var planningResult = planner.Plan(clashZones);

    // 3. Use planningResult.Items for ordered processing:
    foreach (var dto in planningResult.Items)
    {
        if (dto.ShouldSkip) continue; // Skip early
        // Use dto.TargetWidthFt, dto.TargetHeightFt, dto.RequiredDepthFt, dto.RotationAngleDeg
        // to drive sleeve family selection and parameter assignment.
    }

    // 4. Batch log:
    foreach (var line in planningResult.Items.Select(i => i.LogSummary))
    {
        SafeFileLogger.SafeAppendText("planning_debug.log", line + "\n");
    }

    // 5. Metrics (planningResult.HighRiskCount, etc.) can feed performance report.
*/
