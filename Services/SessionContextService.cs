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
            List<string> selectedHostTypes = null,
            bool skipReset = false)
        {
            _logger($"[SessionContext] 🔄 Updating Session Flags for {filterNames?.Count ?? 0} filters, {categories?.Count ?? 0} categories, {selectedHostTypes?.Count ?? 0} host types");

            // STEP 1: RESET — all zones (flag management independent of filters)
            if (!skipReset)
            {
                _logger("[SessionContext] 1. Resetting Flags for all zones...");
                _repository.ResetIsCurrentClashFlag(null, null); // Null args = Reset ALL
            }
            else
            {
                _logger("[SessionContext] 1. Skipping global reset (already handled by caller).");
            }

            // STEP 2: SET CURRENT (Spatial + Host Type)
            // This marks what is "in session" based on the user's 3D view context AND selected host types.
            if (sectionBox != null)
            {
                _logger("[SessionContext] 2. Setting IsCurrentClashFlag based on Section Box + Host Types...");
                int currentCount = _repository.SetIsCurrentClashFlagSpatially(sectionBox, selectedHostTypes);
                _logger($"[SessionContext]    -> Marked {currentCount} zones as Current (In Box + HostType filter).");
            }
            else
            {
                 _logger("[SessionContext] ⚠️ No Section Box active. IsCurrentClashFlag will be 0 for all zones.");
            }

            // STEP 3: SET READY (In Context + Filter Match + Host Type + Unresolved)
            // This is the final gate: must be Spatially Current AND match selected Filters AND Host Types AND be unresolved.
            _logger("[SessionContext] 3. Setting ReadyForPlacement based on Filter Context...");
            int readyCount = _repository.SetReadyForPlacementWithFilterContext(filterNames, categories, selectedHostTypes);
            
            _logger($"[SessionContext]    -> Marked {readyCount} zones as Ready for Placement.");
            return readyCount;
        }

        // Legacy helper removed in favor of optimized repository methods.
        private int ApplySectionBoxContext(List<string> filterNames, List<string> categories, BoundingBoxXYZ sectionBox, List<string> selectedHostTypes = null)
        {
            return 0; // No-op, called by external code if any remains, but internal logic now uses DB methods.
        }
    }
}
