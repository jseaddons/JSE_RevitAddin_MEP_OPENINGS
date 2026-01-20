using System;
using System.Collections.Generic;
using System.Linq;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor;
using JSE_RevitAddin_MEP_OPENINGS.Services.Logging;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.FilterManagement
{
    /// <summary>
    /// Team J: SOLID-compliant filter CRUD service.
    /// Handles filter creation, copying, renaming, deletion, and validation.
    /// 
    /// ✅ PRESERVES ALL LOGIC:
    /// - CreateFilterFromCurrentUIState logic
    /// - Category validation from UI state
    /// - Deep copy logic for filter copying
    /// - Filter validation (category, name)
    /// </summary>
    public class FilterCrudService : IFilterCrudService
    {
        private readonly ILogger _logger;
        
        /// <summary>
        /// Creates a new filter CRUD service.
        /// </summary>
        public FilterCrudService(ILogger logger = null)
        {
            _logger = logger ?? LoggerAdapter.Default;
        }
        
        /// <summary>
        /// Creates a new filter from current UI state.
        /// ✅ PRESERVE: CreateFilterFromCurrentUIState logic
        /// ✅ PRESERVE: Category validation from UI state
        /// ✅ REUSE: FilterUiStateProvider for UI state access
        /// </summary>
        public OpeningFilter CreateFilter(string filterName)
        {
            if (string.IsNullOrWhiteSpace(filterName))
            {
                _logger.Warning("Filter name cannot be null or empty", "FilterCrudService");
                return null;
            }
            
            var filter = new OpeningFilter
            {
                Name = filterName,
                Category = Models.MepCategory.Ducts,
                OpeningType = Models.OpeningType.RectangularSleeves,
                IsEnabled = true,
                LastModified = DateTime.Now,
                ClashZoneStorage = null
            };

            // ✅ PRESERVE: Pull current UI selections via provider
            try
            {
                var cats = FilterUiStateProvider.GetSelectedMepCategoryNames?.Invoke();
                if (cats != null && cats.Count > 0)
                {
                    filter.SelectedMepCategoryNames = new List<string>(cats);
                    filter.SelectedMepCategoryName = cats[0];
                }

                var refs = FilterUiStateProvider.GetSelectedReferenceFiles?.Invoke();
                if (refs != null) filter.SelectedReferenceFiles = new List<string>(refs);

                var hosts = FilterUiStateProvider.GetSelectedHostFiles?.Invoke();
                if (hosts != null) filter.SelectedHostFiles = new List<string>(hosts);

                // ✅ Use UI state from provider
                var hostCategories = FilterUiStateProvider.GetSelectedHostCategories?.Invoke();
                if (hostCategories != null && hostCategories.Count > 0)
                {
                    filter.SelectedHostCategories = new List<string>(hostCategories);
                }
                else
                {
                    // No UI state from provider - initialize empty
                    filter.SelectedHostCategories = new List<string>();
                }
            }
            catch (Exception ex)
            {
                _logger.Warning($"Error getting UI state: {ex.Message}", "FilterCrudService");
            }

            return filter;
        }
        
        /// <summary>
        /// Copies an existing filter.
        /// ✅ PRESERVE: Deep copy logic
        /// ✅ PRESERVE: Category handling
        /// </summary>
        public OpeningFilter CopyFilter(OpeningFilter sourceFilter, string newName)
        {
            if (sourceFilter == null)
            {
                _logger.Warning("Source filter cannot be null", "FilterCrudService");
                return null;
            }
            
            if (string.IsNullOrWhiteSpace(newName))
            {
                _logger.Warning("New filter name cannot be null or empty", "FilterCrudService");
                return null;
            }
            
            var copiedFilter = new OpeningFilter
            {
                Name = newName,
                Category = sourceFilter.Category,
                OpeningType = sourceFilter.OpeningType,
                IsEnabled = sourceFilter.IsEnabled,
                ClashZoneStorage = sourceFilter.ClashZoneStorage,
                LastModified = DateTime.Now,
                // ✅ COPY UI STATE: Copy all UI state properties from source filter
                SelectedMepCategoryNames = sourceFilter.SelectedMepCategoryNames != null ? new List<string>(sourceFilter.SelectedMepCategoryNames) : null,
                SelectedMepCategoryName = sourceFilter.SelectedMepCategoryName,
                SelectedReferenceFiles = sourceFilter.SelectedReferenceFiles != null ? new List<string>(sourceFilter.SelectedReferenceFiles) : null,
                SelectedHostFiles = sourceFilter.SelectedHostFiles != null ? new List<string>(sourceFilter.SelectedHostFiles) : null,
                SelectedHostCategories = sourceFilter.SelectedHostCategories != null ? new List<string>(sourceFilter.SelectedHostCategories) : null,
                OpeningSettings = sourceFilter.OpeningSettings // ✅ CRITICAL: Copy OpeningSettings as well
            };
            
            return copiedFilter;
        }
        
        /// <summary>
        /// Validates filter before save.
        /// ✅ PRESERVE: Category validation
        /// ✅ PRESERVE: Name validation
        /// </summary>
        public (bool isValid, string errorMessage) ValidateFilter(OpeningFilter filter)
        {
            if (filter == null)
            {
                return (false, "Filter cannot be null");
            }
            
            if (string.IsNullOrWhiteSpace(filter.Name))
            {
                return (false, "Filter name cannot be empty");
            }
            
            // ✅ CRITICAL: Validate category from UI state, NOT from filter object (which may have stale/wrong category)
            string categoryDisplay = null;
            if (FilterUiStateProvider.GetSelectedMepCategoryNames != null)
            {
                try
                {
                    var selectedCategories = FilterUiStateProvider.GetSelectedMepCategoryNames.Invoke();
                    if (selectedCategories != null && selectedCategories.Count > 0)
                    {
                        categoryDisplay = MepCategoryConstants.Normalize(selectedCategories[0]);
                    }
                }
                catch (Exception ex)
                {
                    _logger.Warning($"Error getting category from UI state: {ex.Message}", "FilterCrudService");
                }
            }
            
            // Fallback to filter's SelectedMepCategoryNames if UI state not available
            if (string.IsNullOrEmpty(categoryDisplay) && filter.SelectedMepCategoryNames != null && filter.SelectedMepCategoryNames.Count > 0)
            {
                categoryDisplay = MepCategoryConstants.Normalize(filter.SelectedMepCategoryNames[0]);
            }
            
            // ⚠️ CRITICAL: If category is empty, validation fails
            if (string.IsNullOrEmpty(categoryDisplay))
            {
                return (false, "Please select a MEP category (Ducts, Pipes, or Cable Trays) before saving the filter.");
            }
            
            return (true, null);
        }
    }
}

