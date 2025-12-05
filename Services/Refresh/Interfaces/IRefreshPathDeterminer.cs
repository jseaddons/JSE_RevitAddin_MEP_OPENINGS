using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Refresh
{
    /// <summary>
    /// Interface for determining which refresh path strategy to use based on context.
    /// Based on PERSISTENCE_TIMING_FLOW_DIAGRAM.md:
    /// - PATH 1: Replay mode (Adopt OFF, filter exists)
    /// - PATH 2: Fresh mode (first time adding filter)
    /// - PATH 3: Non-fresh mode (filter exists, refresh again)
    /// </summary>
    public interface IRefreshPathDeterminer
    {
        /// <summary>
        /// Determine refresh path strategy based on:
        /// 1. enableThreePointValidation (Adopt to Document checkbox)
        /// 2. FileCombo existence (checked FIRST before determining path)
        /// 3. Whether filter exists (has existing clash zones)
        /// </summary>
        /// <param name="context">Refresh context with UI selections and settings</param>
        /// <returns>Appropriate path strategy (Path1/Path2/Path3)</returns>
        IRefreshPathStrategy DeterminePath(RefreshContext context);
        
        /// <summary>
        /// Determine placement path (checks FileCombo existence in database).
        /// FileCombo existence is checked FIRST (before IsFilterComboNew flag) 
        /// because FileCombos are populated when clash zones are saved.
        /// </summary>
        /// <param name="context">Refresh context</param>
        /// <returns>True if replay path, false if fresh/non-fresh path</returns>
        bool DeterminePlacementPath(RefreshContext context);
        
        /// <summary>
        /// Update IsFilterComboNew flag based on FileCombos in database.
        /// Checks if FileCombos exist for current UI state (filter + files + category).
        /// Sets IsFilterComboNew = true if FileCombos don't exist (fresh combo).
        /// Sets IsFilterComboNew = false if FileCombos exist (used combo).
        /// </summary>
        /// <param name="context">Refresh context</param>
        void UpdateIsFilterComboNewFlagBasedOnFileCombos(RefreshContext context);
    }
}
