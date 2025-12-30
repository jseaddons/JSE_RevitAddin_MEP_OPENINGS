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
            BoundingBoxXYZ sectionBox)
        {
            _logger($"[SessionContext] 🔄 Updating Session Flags for {filterNames?.Count ?? 0} filters, {categories?.Count ?? 0} categories");

            // STEP 1: RESET
            _logger("[SessionContext] 1. Resetting IsCurrentClashFlag for all zones...");
            _repository.ResetIsCurrentClashFlag();

            // STEP 2: SET CURRENT (Spatial + Filter)
            if (sectionBox != null)
            {
                _logger("[SessionContext] 2. Setting IsCurrentClashFlag based on Section Box...");
                int currentCount = ApplySectionBoxContext(filterNames, categories, sectionBox);
                _logger($"[SessionContext]    -> Marked {currentCount} zones as Current (In Box + Filter match).");
            }
            else
            {
                 _logger("[SessionContext] ⚠️ No Section Box active. IsCurrentClashFlag will be 0 for all zones (Safety Default).");
            }

            // STEP 3: SET READY (Current + Unresolved)
            _logger("[SessionContext] 3. Setting ReadyForPlacement for Unresolved Current zones...");
            int readyCount = _repository.SetReadyForPlacementForUnresolvedZonesInSectionBox(null, null, null); 
            // Note: We pass nulls because the repository method now relies purely on IsCurrentClashFlag=1
            
            _logger($"[SessionContext]    -> Marked {readyCount} zones as Ready for Placement.");
            return readyCount;
        }

        private int ApplySectionBoxContext(List<string> filterNames, List<string> categories, BoundingBoxXYZ sectionBox)
        {
            // Business Logic: Which zones are "Current"?
            // They must match the Filter/Category AND be inside the Section Box.
            
            // This logic was previously inside the Repository, which violated SOLID (SRP).
            // We moved it here to orchestrate.

            // 1. Get Candidate IDs (Filter + Category) directly from DB.
            //    This is efficient as it uses the existing optimized query.
            var candidates = _repository.GetClashZonesByFilterAndCategory(filterNames, categories);
            
            if (candidates.Count == 0) return 0;

            // 2. Filter Spatially (In Memory Check)
            //    Since we have the candidates, we check against the box.
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

            // 3. Commit to DB
            if (zonesToMark.Count > 0)
            {
                _repository.BulkSetIsCurrentClashFlag(zonesToMark, true);
            }

            return zonesToMark.Count;
        }
    }
}
