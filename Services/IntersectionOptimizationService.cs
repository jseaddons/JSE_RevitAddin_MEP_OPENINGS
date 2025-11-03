using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Service for optimizing intersection detection by skipping geometry checks for known valid element pairs.
    /// Follows OOP principles: Single Responsibility (optimization logic only), Encapsulation (private state).
    /// </summary>
    public class IntersectionOptimizationService
    {
        private readonly HashSet<(int mepId, int structuralId)> _knownValidPairs;
        private readonly bool _skipKnownPairsGeometryCheck;
        private readonly SettingsService _settingsService;

        /// <summary>
        /// Creates an optimization service that determines whether to skip geometry checks for known pairs.
        /// </summary>
        /// <param name="existingClashZones">Existing clash zones from XML</param>
        /// <param name="settingsService">Settings service to check if 3-point validation is enabled</param>
        public IntersectionOptimizationService(
            ClashZoneStorage existingClashZones,
            SettingsService settingsService = null)
        {
            _settingsService = settingsService ?? new SettingsService();
            _knownValidPairs = new HashSet<(int, int)>();
            _skipKnownPairsGeometryCheck = false;

            if (existingClashZones?.ClashZones != null && existingClashZones.ClashZones.Count > 0)
            {
                var settings = _settingsService.LoadSettings();
                bool enableThreePointValidation = settings?.EnableThreePointValidation ?? true;

                if (!enableThreePointValidation)
                {
                    // ✅ OPTIMIZATION: Build set of known valid (MEP, Structural) pairs
                    // When 3-point validation is disabled, user trusts model unchanged
                    // Skip expensive geometry intersection for these pairs - just verify element existence
                    foreach (var cz in existingClashZones.ClashZones)
                    {
                        if (cz != null && (cz.IsResolved || cz.IsClusterResolved))
                        {
                            int mepId = cz.MepElementId?.IntegerValue ?? cz.MepElementIdValue;
                            int structuralId = cz.StructuralElementId?.IntegerValue ?? cz.StructuralElementIdValue;

                            if (mepId > 0 && structuralId > 0)
                            {
                                _knownValidPairs.Add((mepId, structuralId));
                            }
                        }
                    }

                    _skipKnownPairsGeometryCheck = _knownValidPairs.Count > 0;
                }
            }
        }

        /// <summary>
        /// Gets the set of known valid element pairs (MEP ID, Structural ID)
        /// </summary>
        public HashSet<(int mepId, int structuralId)> KnownValidPairs => _knownValidPairs;

        /// <summary>
        /// Returns true if geometry checks should be skipped for known pairs
        /// </summary>
        public bool ShouldSkipKnownPairsGeometryCheck => _skipKnownPairsGeometryCheck;

        /// <summary>
        /// Checks if a (MEP Element, Structural Element) pair is known and valid.
        /// If true and optimization is enabled, geometry intersection can be skipped.
        /// </summary>
        public bool IsKnownValidPair(Element mepElement, Element structuralElement)
        {
            if (!_skipKnownPairsGeometryCheck || mepElement == null || structuralElement == null)
                return false;

            int mepId = mepElement.Id.IntegerValue;
            int structuralId = structuralElement.Id.IntegerValue;
            return _knownValidPairs.Contains((mepId, structuralId));
        }

        /// <summary>
        /// Checks if a (MEP Element ID, Structural Element ID) pair is known and valid.
        /// </summary>
        public bool IsKnownValidPair(int mepElementId, int structuralElementId)
        {
            if (!_skipKnownPairsGeometryCheck || mepElementId <= 0 || structuralElementId <= 0)
                return false;

            return _knownValidPairs.Contains((mepElementId, structuralElementId));
        }

        /// <summary>
        /// Gets optimization status for logging
        /// </summary>
        public string GetOptimizationStatus()
        {
            if (_skipKnownPairsGeometryCheck)
            {
                return $"Optimization ENABLED: {_knownValidPairs.Count} known valid pairs - geometry checks will be skipped";
            }
            else
            {
                return "Optimization DISABLED: Full intersection detection will run (3-point validation enabled or no known pairs)";
            }
        }
    }
}

