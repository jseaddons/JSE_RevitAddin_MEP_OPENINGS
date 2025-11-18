using System;
using System.Collections.Generic;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Lightweight static provider to bridge UI state to FilterManagementService without tight coupling.
    /// Callers (e.g., EmergencyMainDialog) should register delegates on startup.
    /// </summary>
    public static class FilterUiStateProvider
    {
        public static Func<List<string>> GetSelectedMepCategoryNames { get; set; }
        public static Func<List<string>> GetSelectedReferenceFiles { get; set; }
        public static Func<List<string>> GetSelectedHostFiles { get; set; }
        public static Func<List<string>> GetSelectedHostCategories { get; set; }
        public static Func<List<string>> GetSelectedFilterItems { get; set; }
        public static Func<string, Dictionary<string, double>> GetClearanceSettings { get; set; }

        public static Action<OpeningFilter> ApplyFilterToUi { get; set; }
    }
}


