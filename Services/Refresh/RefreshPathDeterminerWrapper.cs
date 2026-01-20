using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Refresh;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Refresh
{
    /// <summary>
    /// Wrapper for the static RefreshPathDeterminer to support dependency injection.
    /// Implements IRefreshPathDeterminer.
    /// </summary>
    public class RefreshPathDeterminerWrapper : IRefreshPathDeterminer
    {
        /// <summary>
        /// Determine refresh path strategy based on context.
        /// Delegates to static RefreshPathDeterminer.DeterminePath.
        /// </summary>
        public IRefreshPathStrategy DeterminePath(RefreshContext context)
        {
            if (context == null) throw new ArgumentNullException(nameof(context));
            
            // Extract enableThreePointValidation from context settings
            return RefreshPathDeterminer.DeterminePath(context, context.EnableThreePointValidation);
        }

        /// <summary>
        /// Determine placement path (checks FileCombo existence).
        /// Returns true if Replay path, false if Sizing/Fresh path.
        /// Aggregates results across all selected filters and categories.
        /// </summary>
        public bool DeterminePlacementPath(RefreshContext context)
        {
            if (context == null) throw new ArgumentNullException(nameof(context));
            
            // If any combination requires Sizing (Fresh), then the overall result is NOT Replay.
            // Only return true (Replay) if ALL combinations are Replay.
            
            if (context.SelectedFilterNames == null || context.SelectedMepCategories == null)
                return false; // Safe default
                
            foreach (var filterName in context.SelectedFilterNames)
            {
                foreach (var category in context.SelectedMepCategories)
                {
                    var path = RefreshPathDeterminer.DeterminePlacementPath(
                        context.Document,
                        filterName,
                        category,
                        "", // No log prefix needed here
                        context.ClearanceSettings);
                        
                    if (path != SleevePlacementPath.Replay)
                    {
                        // Found a fresh/sizing path, so we cannot do a full replay
                        return false;
                    }
                }
            }
            
            // All combinations are Replay
            return true;
        }

        /// <summary>
        /// Update IsFilterComboNew flag based on FileCombos in database.
        /// Delegates to static RefreshPathDeterminer.
        /// </summary>
        public void UpdateIsFilterComboNewFlagBasedOnFileCombos(RefreshContext context)
        {
            if (context == null) throw new ArgumentNullException(nameof(context));
            
            RefreshPathDeterminer.UpdateIsFilterComboNewFlagBasedOnFileCombos(context);
        }
    }
}
