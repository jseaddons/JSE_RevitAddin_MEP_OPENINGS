using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Models;


namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Service responsible for managing the "Session Context" of clash zones.
    /// This defines which zones are "Current" based on UI filters and 3D Section Box.
    /// Encapsulates the Business Logic of flag management, keeping Repository pure.
    /// </summary>
    public class SessionContextService
    {
        private readonly ClashZoneRepository _repository;
        private readonly Action<string> _logger;

        public SessionContextService(ClashZoneRepository repository, Action<string> logger = null)
        {
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
            _logger = logger ?? ((msg) => { });
        }


        public int UpdateSessionFlags(
            List<string> filterNames, 
            List<string> categories, 
            BoundingBoxXYZ sectionBox,
            List<string> selectedHostTypes = null)
        {
            _logger($"[SessionContext] 🔄 Updating Session Flags for {filterNames?.Count ?? 0} filters, {categories?.Count ?? 0} categories, {selectedHostTypes?.Count ?? 0} host types");

            // STEP 0: Clear ReadyForPlacement for selected filters/categories so only zones we mark as Current get it set again.
            // Without this, zones from a previous run (e.g. floor when user didn't select floor) keep ReadyForPlacement=1 and get placed.
            if (filterNames != null && filterNames.Count > 0 && categories != null && categories.Count > 0)
            {
                int resetReady = _repository.BulkResetReadyForPlacementFlagsForFilters(filterNames, categories);
                _logger($"[SessionContext] 0. Cleared ReadyForPlacement for {resetReady} zones in selected scope.");
            }

            // STEP 1: RESET — all zones (flag management independent of filters; e.g. Electrical can have Ducts)
            _logger("[SessionContext] 1. Resetting IsCurrentClashFlag for all zones...");
            _repository.ResetIsCurrentClashFlag();

            // STEP 2: SET CURRENT (Spatial + Filter + Host type)
            if (sectionBox != null)
            {
                _logger("[SessionContext] 2. Setting IsCurrentClashFlag based on Section Box (and selected host types)...");
                int currentCount = ApplySectionBoxContext(filterNames, categories, sectionBox, selectedHostTypes);
                _logger($"[SessionContext]    -> Marked {currentCount} zones as Current (In Box + Filter + Host match).");
            }
            else
            {
                 _logger("[SessionContext] ⚠️ No Section Box active. IsCurrentClashFlag will be 0 for all zones (Safety Default).");
            }

            // STEP 3: SET READY (Current + Unresolved) — flag management independent of filters
            _logger("[SessionContext] 3. Setting ReadyForPlacement for Unresolved Current zones...");
            int readyCount = _repository.SetReadyForPlacementForUnresolvedZonesInSectionBox(null, null, null);
            
            _logger($"[SessionContext]    -> Marked {readyCount} zones as Ready for Placement.");
            return readyCount;
        }

        private int ApplySectionBoxContext(List<string> filterNames, List<string> categories, BoundingBoxXYZ sectionBox, List<string> selectedHostTypes = null)
        {
            // Business Logic: Which zones are "Current"?
            // They must match the Filter/Category AND be inside the Section Box AND (if host types selected) match host type.

            // 1. Get Candidate IDs (Filter + Category) directly from DB.
            var candidates = _repository.GetClashZonesByFilterAndCategory(filterNames, categories);
            if (candidates.Count == 0) return 0;

            // 2. Filter by selected host types (e.g. Walls, Floors, Structural Framing). If user didn't select Floor, exclude floor sleeves.
            if (selectedHostTypes != null && selectedHostTypes.Count > 0)
            {
                var allowed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var h in selectedHostTypes)
                {
                    if (string.IsNullOrWhiteSpace(h)) continue;
                    allowed.Add(h.Trim());
                    // Normalize: "Floors" <-> "Floor", "Walls" <-> "Wall"
                    if (h.Trim().Equals("Floors", StringComparison.OrdinalIgnoreCase)) allowed.Add("Floor");
                    if (h.Trim().Equals("Floor", StringComparison.OrdinalIgnoreCase)) allowed.Add("Floors");
                    if (h.Trim().Equals("Walls", StringComparison.OrdinalIgnoreCase)) allowed.Add("Wall");
                    if (h.Trim().Equals("Wall", StringComparison.OrdinalIgnoreCase)) allowed.Add("Walls");
                }
                candidates = candidates.Where(z => !string.IsNullOrEmpty(z.StructuralElementType) && allowed.Contains(z.StructuralElementType.Trim())).ToList();
                if (candidates.Count == 0) return 0;
            }

            // 3. Filter Spatially (In Memory Check)
            var zonesToMark = new List<Guid>();
            foreach (var zone in candidates)
            {
                var pt = zone.IntersectionPoint;
                if (pt == null && (Math.Abs(zone.IntersectionPointX) > 1e-9 || Math.Abs(zone.IntersectionPointY) > 1e-9))
                {
                    pt = new XYZ(zone.IntersectionPointX, zone.IntersectionPointY, zone.IntersectionPointZ);
                }

                if (pt != null && JSE_RevitAddin_MEP_OPENINGS.Helpers.SectionBoxHelper.IsPointInBoundingBox(pt, sectionBox))
                {
                    zonesToMark.Add(zone.Id);
                }
            }

            // 4. Commit to DB using high-performance Temp Table merge
            if (zonesToMark.Count > 0)
            {
                _repository.BulkSetIsCurrentClashFlagTempTable(zonesToMark, true);
            }

            return zonesToMark.Count;
        }
    }
}
