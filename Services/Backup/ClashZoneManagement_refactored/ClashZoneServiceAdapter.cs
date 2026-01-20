using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Logging;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.ClashZoneManagement
{
    /// <summary>
    /// Adapter that wraps the legacy ClashZoneService so the refactored refresh pipeline
    /// can continue to depend on the IClashZoneService abstraction.
    /// </summary>
    public class ClashZoneServiceAdapter : IClashZoneService
    {
        private readonly bool _useRefactored;
        private readonly IClashZoneService? _refactoredService;
        private readonly Services.ClashZoneService? _legacyService;
        private readonly ILogger _logger;

        public ClashZoneServiceAdapter(
            Services.ClashZoneService legacyService,
            ILogger logger = null,
            IClashZoneService refactoredService = null)
        {
            _logger = logger ?? LoggerAdapter.Default;
            _useRefactored = OptimizationFlags.UseRefactoredClashZoneFlagServices;
            _legacyService = legacyService;
            _refactoredService = refactoredService;

            if (_useRefactored && _refactoredService == null)
            {
                _logger.Warning("ClashZoneServiceAdapter created in refactored mode without a refactored service instance. Falling back to legacy service.", "ClashZoneServiceAdapter");
                _useRefactored = false;
            }
        }

        public int CleanupInvalidClashZones(Document document)
        {
            if (_useRefactored && _refactoredService != null)
            {
                return _refactoredService.CleanupInvalidClashZones(document);
            }

            if (_legacyService == null)
                throw new InvalidOperationException("Legacy clash zone service is not available.");

            return _legacyService.CleanupInvalidClashZones(document);
        }

        public List<ClashZone> FilterClashZonesByCurrentSelection(
            List<string> selectedReferenceFiles,
            Dictionary<string, double> currentClearanceSettings,
            string currentPrefix,
            Document document)
        {
            if (_useRefactored && _refactoredService != null)
            {
                return _refactoredService.FilterClashZonesByCurrentSelection(
                    selectedReferenceFiles,
                    currentClearanceSettings,
                    currentPrefix,
                    document);
            }

            if (_legacyService == null)
                throw new InvalidOperationException("Legacy clash zone service is not available.");

            return _legacyService.FilterClashZonesByCurrentSelection(
                selectedReferenceFiles,
                currentClearanceSettings,
                currentPrefix,
                document);
        }

        public List<ClashZone> DetectNewClashZones(
            List<(Element, Element, BoundingBoxXYZ, XYZ)> currentIntersections,
            Document document,
            Dictionary<string, double> clearanceSettings = null,
            List<string> selectedCategories = null)
        {
            if (_useRefactored && _refactoredService != null)
            {
                return _refactoredService.DetectNewClashZones(
                    currentIntersections,
                    document,
                    clearanceSettings,
                    selectedCategories);
            }

            if (_legacyService == null)
                throw new InvalidOperationException("Legacy clash zone service is not available.");

            return _legacyService.DetectNewClashZones(
                currentIntersections,
                document,
                clearanceSettings,
                selectedCategories);
        }
    }
}

