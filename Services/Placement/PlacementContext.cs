using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Placement
{
    /// <summary>
    /// Immutable context passed through the placement pipeline stages.
    /// Uses builder-style With* methods to create new instances with updated data.
    /// Follows SOLID principles: immutability reduces side effects, improves testability.
    /// </summary>
    public sealed class PlacementContext
    {
        /// <summary>
        /// Unique identifier for tracing this placement operation through logs.
        /// </summary>
        public string CorrelationId { get; }

        /// <summary>
        /// Revit document for element operations.
        /// </summary>
        public Document Doc { get; }

        /// <summary>
        /// Original input clash zones for placement.
        /// </summary>
        public IReadOnlyList<ClashZone> InputZones { get; }

        /// <summary>
        /// Configuration driving optimization and feature flags.
        /// </summary>
        public IOptimizationConfig Config { get; }

        /// <summary>
        /// Zones remaining after filtering stage.
        /// </summary>
        public IReadOnlyList<ClashZone> FilteredZones { get; }

        /// <summary>
        /// Calculated dimensions per zone (zoneId -> (width, height, diameter, isCircular)).
        /// </summary>
        public IReadOnlyDictionary<Guid, (double width, double height, double diameter, bool circular)> Dimensions { get; }

        /// <summary>
        /// Successfully placed sleeve instances.
        /// </summary>
        public IReadOnlyList<FamilyInstance> PlacedInstances { get; }

        /// <summary>
        /// Accumulated errors from stages (non-critical failures for fail-safe semantics).
        /// </summary>
        public IReadOnlyList<string> Errors { get; }

        /// <summary>
        /// Performance and diagnostic metrics (counts, timings, flags).
        /// </summary>
        public IReadOnlyDictionary<string, object> Metrics { get; }

        public PlacementContext(
            Document doc,
            IReadOnlyList<ClashZone> zones,
            IOptimizationConfig config,
            IReadOnlyList<ClashZone> filtered = null,
            IReadOnlyDictionary<Guid, (double, double, double, bool)> dims = null,
            IReadOnlyList<FamilyInstance> placed = null,
            IReadOnlyList<string> errors = null,
            IReadOnlyDictionary<string, object> metrics = null,
            string correlationId = null)
        {
            CorrelationId = correlationId ?? Guid.NewGuid().ToString();
            Doc = doc ?? throw new ArgumentNullException(nameof(doc));
            InputZones = zones ?? throw new ArgumentNullException(nameof(zones));
            Config = config ?? throw new ArgumentNullException(nameof(config));
            FilteredZones = filtered ?? Array.Empty<ClashZone>();
            Dimensions = dims ?? new Dictionary<Guid, (double, double, double, bool)>();
            PlacedInstances = placed ?? Array.Empty<FamilyInstance>();
            Errors = errors ?? Array.Empty<string>();
            Metrics = metrics ?? new Dictionary<string, object>();
        }

        /// <summary>
        /// Create new context with updated filtered zones.
        /// </summary>
        public PlacementContext WithFilteredZones(IReadOnlyList<ClashZone> filtered) =>
            new PlacementContext(Doc, InputZones, Config, filtered, Dimensions, PlacedInstances, Errors, Metrics, CorrelationId);

        /// <summary>
        /// Create new context with additional dimension entry.
        /// </summary>
        public PlacementContext WithDimension(Guid zoneId, (double w, double h, double d, bool c) dim)
        {
            // Create new dictionary with existing dimensions plus new entry
            var dict = new Dictionary<Guid, (double, double, double, bool)>();
            foreach (var kvp in Dimensions)
            {
                dict[kvp.Key] = kvp.Value;
            }
            dict[zoneId] = dim;
            return new PlacementContext(Doc, InputZones, Config, FilteredZones, dict, PlacedInstances, Errors, Metrics, CorrelationId);
        }

        /// <summary>
        /// Create new context with all dimensions replaced.
        /// </summary>
        public PlacementContext WithDimensions(IReadOnlyDictionary<Guid, (double, double, double, bool)> dims) =>
            new PlacementContext(Doc, InputZones, Config, FilteredZones, dims, PlacedInstances, Errors, Metrics, CorrelationId);

        /// <summary>
        /// Create new context with additional placed instance.
        /// </summary>
        public PlacementContext WithPlaced(FamilyInstance inst)
        {
            var list = PlacedInstances.ToList();
            list.Add(inst);
            return new PlacementContext(Doc, InputZones, Config, FilteredZones, Dimensions, list, Errors, Metrics, CorrelationId);
        }

        /// <summary>
        /// Create new context with all placed instances replaced.
        /// </summary>
        public PlacementContext WithPlacedInstances(IReadOnlyList<FamilyInstance> placed) =>
            new PlacementContext(Doc, InputZones, Config, FilteredZones, Dimensions, placed, Errors, Metrics, CorrelationId);

        /// <summary>
        /// Create new context with additional error message.
        /// </summary>
        public PlacementContext WithError(string error)
        {
            var list = Errors.ToList();
            list.Add(error);
            return new PlacementContext(Doc, InputZones, Config, FilteredZones, Dimensions, PlacedInstances, list, Metrics, CorrelationId);
        }

        /// <summary>
        /// Create new context with multiple errors added.
        /// </summary>
        public PlacementContext WithErrors(IEnumerable<string> errors)
        {
            var list = Errors.ToList();
            list.AddRange(errors);
            return new PlacementContext(Doc, InputZones, Config, FilteredZones, Dimensions, PlacedInstances, list, Metrics, CorrelationId);
        }

        /// <summary>
        /// Create new context with updated metric.
        /// </summary>
        public PlacementContext WithMetric(string key, object value)
        {
            var dict = new Dictionary<string, object>();
            foreach (var kvp in Metrics)
            {
                dict[kvp.Key] = kvp.Value;
            }
            dict[key] = value;
            return new PlacementContext(Doc, InputZones, Config, FilteredZones, Dimensions, PlacedInstances, Errors, dict, CorrelationId);
        }

        /// <summary>
        /// Create new context with multiple metrics added/updated.
        /// </summary>
        public PlacementContext WithMetrics(IEnumerable<KeyValuePair<string, object>> metrics)
        {
            var dict = new Dictionary<string, object>();
            foreach (var kvp in Metrics)
            {
                dict[kvp.Key] = kvp.Value;
            }
            foreach (var kvp in metrics)
            {
                dict[kvp.Key] = kvp.Value;
            }
            return new PlacementContext(Doc, InputZones, Config, FilteredZones, Dimensions, PlacedInstances, Errors, dict, CorrelationId);
        }
    }
}
