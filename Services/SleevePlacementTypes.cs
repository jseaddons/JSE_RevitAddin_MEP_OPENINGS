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
    /// ✅ Extracted from SleevePlacementCoordinator.cs for use while that file is excluded from build.
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
    /// ✅ Extracted from SleevePlacementCoordinator.cs for use while that file is excluded from build.
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
        public List<ClashZone> ClashZones { get; }
        public string Category { get; }
        public string FilterName { get; }
        public OpeningConditions Conditions { get; }
        public ISleevePlacementStrategy Strategy { get; }
        public Dictionary<string, double> ClearanceSettings { get; }
        public SleevePlacementPath RequestedPath { get; }
        public FlagManager FlagManager { get; }
        public string UiOverridesFingerprint { get; }
        public string GlobalConfigurationFingerprint { get; }
    }

    /// <summary>
    /// Output from sleeve placement pipeline. Encapsulates success/failure state,
    /// placed sleeves, and any diagnostics.
    /// ✅ Extracted from SleevePlacementCoordinator.cs for use while that file is excluded from build.
    /// </summary>
    public sealed class SleevePlacementResult
    {
        public static SleevePlacementResult Success(
            List<FamilyInstance> placedSleeves,
            List<ClashZone> updatedClashZones,
            SleevePlacementPath pathUsed,
            Dictionary<string, object> metadata = null)
        {
            return new SleevePlacementResult
            {
                IsSuccess = true,
                PlacedSleeves = placedSleeves ?? new List<FamilyInstance>(),
                UpdatedClashZones = updatedClashZones ?? new List<ClashZone>(),
                PathUsed = pathUsed,
                Metadata = metadata ?? new Dictionary<string, object>()
            };
        }

        public static SleevePlacementResult Failure(string errorMessage, Exception exception = null)
        {
            return new SleevePlacementResult
            {
                IsSuccess = false,
                ErrorMessage = errorMessage,
                Exception = exception,
                PlacedSleeves = new List<FamilyInstance>(),
                UpdatedClashZones = new List<ClashZone>(),
                Metadata = new Dictionary<string, object>()
            };
        }

        public bool IsSuccess { get; private set; }
        public List<FamilyInstance> PlacedSleeves { get; private set; }
        public List<ClashZone> UpdatedClashZones { get; private set; }
        public SleevePlacementPath PathUsed { get; private set; }
        public Dictionary<string, object> Metadata { get; private set; }
        public string ErrorMessage { get; private set; }
        public Exception Exception { get; private set; }
    }

    /// <summary>
    /// Path handler interface for individual placement strategies (Replay, Sizing, Detection).
    /// ✅ Extracted from SleevePlacementCoordinator.cs for use while that file is excluded from build.
    /// </summary>
    public interface ISleevePlacementPathHandler
    {
        SleevePlacementPath Path { get; }
        SleevePlacementResult Execute(SleevePlacementRequest request);
    }
}
