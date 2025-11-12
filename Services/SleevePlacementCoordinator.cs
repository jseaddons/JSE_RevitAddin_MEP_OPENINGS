using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Strategies;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Identifies the placement pipeline path that should run for the current request.
    /// </summary>
    public enum SleevePlacementPath
    {
        Replay = 0,
        Sizing = 1,
        Detection = 2,
    }

    /// <summary>
    /// Input payload for sleeve placement. Acts as the shared context between the
    /// coordinator and individual path handlers.
    /// </summary>
    public sealed class SleevePlacementRequest
    {
        public SleevePlacementRequest(
            Document document,
            IEnumerable<ClashZone> clashZones,
            string category,
            string filterName,
            OpeningConditions conditions,
            ISleevePlacementStrategy strategy,
            IDictionary<string, double> clearanceSettings,
            SleevePlacementPath requestedPath,
            FlagManager flagManager = null,
            string uiOverridesFingerprint = null,
            string globalConfigurationFingerprint = null)
        {
            Document = document ?? throw new ArgumentNullException(nameof(document));
            ClashZones = (clashZones ?? Enumerable.Empty<ClashZone>()).ToList();
            Category = category ?? throw new ArgumentNullException(nameof(category));
            FilterName = filterName;
            Conditions = conditions ?? new OpeningConditions();
            Strategy = strategy ?? throw new ArgumentNullException(nameof(strategy));
            ClearanceSettings = clearanceSettings != null
                ? new Dictionary<string, double>(clearanceSettings)
                : new Dictionary<string, double>();
            RequestedPath = requestedPath;
            FlagManager = flagManager;
            UiOverridesFingerprint = uiOverridesFingerprint;
            GlobalConfigurationFingerprint = globalConfigurationFingerprint;
        }

        public Document Document { get; }
        public IReadOnlyList<ClashZone> ClashZones { get; }
        public string Category { get; }
        public string FilterName { get; }
        public OpeningConditions Conditions { get; }
        public ISleevePlacementStrategy Strategy { get; }
        public IReadOnlyDictionary<string, double> ClearanceSettings { get; }
        public SleevePlacementPath RequestedPath { get; }
        public FlagManager FlagManager { get; }
        public string UiOverridesFingerprint { get; }
        public string GlobalConfigurationFingerprint { get; }

        public SleevePlacementRequest WithPath(SleevePlacementPath path)
        {
            var clearanceCopy = ClearanceSettings != null
                ? ClearanceSettings.ToDictionary(kv => kv.Key, kv => kv.Value)
                : new Dictionary<string, double>();

            return new SleevePlacementRequest(
                Document,
                ClashZones,
                Category,
                FilterName,
                Conditions,
                Strategy,
                clearanceCopy,
                path,
                FlagManager,
                UiOverridesFingerprint,
                GlobalConfigurationFingerprint);
        }
    }

    /// <summary>
    /// Placement execution summary returned by path handlers.
    /// </summary>
    public sealed class SleevePlacementResult
    {
        private SleevePlacementResult(int placed, int skipped, int errors, SleevePlacementPath path)
        {
            PlacedCount = placed;
            SkippedCount = skipped;
            ErrorCount = errors;
            Path = path;
        }

        public int PlacedCount { get; }
        public int SkippedCount { get; }
        public int ErrorCount { get; }
        public SleevePlacementPath Path { get; }

        public static SleevePlacementResult FromCounts(
            int placed,
            int skipped,
            int errors,
            SleevePlacementPath path) =>
            new SleevePlacementResult(placed, skipped, errors, path);
    }

    /// <summary>
    /// Contract implemented by each placement path handler.
    /// </summary>
    public interface ISleevePlacementPathHandler
    {
        SleevePlacementPath Path { get; }
        SleevePlacementResult Execute(SleevePlacementRequest request);
    }

    /// <summary>
    /// Coordinates placement execution by delegating to the handler registered for a requested path.
    /// </summary>
    public sealed class SleevePlacementCoordinator
    {
        private readonly IReadOnlyDictionary<SleevePlacementPath, ISleevePlacementPathHandler> _handlers;

        public SleevePlacementCoordinator(IEnumerable<ISleevePlacementPathHandler> handlers)
        {
            if (handlers == null) throw new ArgumentNullException(nameof(handlers));

            var handlerMap = new Dictionary<SleevePlacementPath, ISleevePlacementPathHandler>();
            foreach (var handler in handlers)
            {
                if (handler == null) continue;
                handlerMap[handler.Path] = handler;
            }

            if (!handlerMap.ContainsKey(SleevePlacementPath.Replay))
            {
                throw new ArgumentException(
                    "SleevePlacementCoordinator requires at least a replay handler.",
                    nameof(handlers));
            }

            _handlers = handlerMap;
        }

        public SleevePlacementResult Execute(SleevePlacementRequest request)
        {
            if (request == null) throw new ArgumentNullException(nameof(request));

            if (!_handlers.TryGetValue(request.RequestedPath, out var handler))
            {
                throw new InvalidOperationException(
                    $"No placement handler registered for path '{request.RequestedPath}'.");
            }

            // Make the chosen path visible in diagnostics.
            try
            {
                var logPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
                System.IO.File.AppendAllText(
                    logPath,
                    $"[{DateTime.Now:HH:mm:ss}] [PLACEMENT-PATH] Selected={request.RequestedPath}, Filter='{request.FilterName}', Category='{request.Category}'\n");
            }
            catch
            {
                // Swallow logging errors – placement should continue even if diagnostics fail.
            }

            return handler.Execute(request);
        }

        public static SleevePlacementCoordinator CreateDefault()
        {
            var handlers = new ISleevePlacementPathHandler[]
            {
                new SleevePlacementReplayService(),
                new SleeveSizingService(),
                new SleeveDetectionService(),
            };

            return new SleevePlacementCoordinator(handlers);
        }
    }

    /// <summary>
    /// Path handler that reuses the persisted snapshot without recomputing geometry.
    /// </summary>
    internal sealed class SleevePlacementReplayService : ISleevePlacementPathHandler
    {
        public SleevePlacementPath Path => SleevePlacementPath.Replay;

        public SleevePlacementResult Execute(SleevePlacementRequest request)
        {
            if (request == null) throw new ArgumentNullException(nameof(request));

            var clearanceSettings = request.ClearanceSettings != null
                ? request.ClearanceSettings.ToDictionary(kv => kv.Key, kv => kv.Value)
                : new Dictionary<string, double>();

            var placerService = new UniversalSleevePlacerService(
                request.Document,
                request.Conditions,
                request.Strategy,
                clearanceSettings,
                request.FilterName,
                request.FlagManager,
                isReplayPath: request.RequestedPath == SleevePlacementPath.Replay);

            var placementOutcome = placerService.PlaceAllSleevesInTransaction(
                request.ClashZones?.ToList() ?? new List<ClashZone>());

            return SleevePlacementResult.FromCounts(
                placementOutcome.PlacedCount,
                placementOutcome.SkippedCount,
                placementOutcome.ErrorCount,
                SleevePlacementPath.Replay);
        }
    }

    /// <summary>
    /// Placeholder for the sizing path. For now it reuses replay logic until the
    /// recalculation workflow is wired in.
    /// </summary>
    internal sealed class SleeveSizingService : ISleevePlacementPathHandler
    {
        private readonly SleevePlacementReplayService _replayService = new SleevePlacementReplayService();

        public SleevePlacementPath Path => SleevePlacementPath.Sizing;

        public SleevePlacementResult Execute(SleevePlacementRequest request)
        {
            if (request == null) throw new ArgumentNullException(nameof(request));

            // Sizing recalculation will be implemented in a follow-up task. Until then,
            // fall back to replay so behaviour remains unchanged.
            try
            {
                var logPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
                System.IO.File.AppendAllText(
                    logPath,
                    $"[{DateTime.Now:HH:mm:ss}] [PLACEMENT-PATH] Sizing path currently delegates to replay (placeholder).\n");
            }
            catch
            {
                // Ignore logging failures
            }

            return _replayService.Execute(request.WithPath(SleevePlacementPath.Replay));
        }
    }

    /// <summary>
    /// Placeholder for the detection path. Will be integrated with the refresh/detection
    /// pipeline in a subsequent iteration.
    /// </summary>
    internal sealed class SleeveDetectionService : ISleevePlacementPathHandler
    {
        public SleevePlacementPath Path => SleevePlacementPath.Detection;

        public SleevePlacementResult Execute(SleevePlacementRequest request)
        {
            throw new NotSupportedException(
                "Sleeve detection placement path is not yet integrated with the coordinator.");
        }
    }
}

