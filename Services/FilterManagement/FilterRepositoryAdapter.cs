using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor;
using JSE_RevitAddin_MEP_OPENINGS.Services.Logging;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.FilterManagement
{
    /// <summary>
    /// Team I: Adapter for FilterRepository to implement IFilterRepository.
    /// SOLID: Adapter pattern for dependency inversion.
    /// DATABASE-ONLY: No XML support.
    /// </summary>
    public class FilterRepositoryAdapter : IFilterRepository
    {
        private readonly Document _document;
        private readonly ILogger _logger;
        
        public FilterRepositoryAdapter(Document document, ILogger logger = null)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _logger = logger ?? LoggerAdapter.Default;
        }
        
        public int EnsureFilter(string filterName, string category)
        {
            int filterId = -1;
            UseFilterRepository(repo => 
            {
                filterId = repo.EnsureFilter(filterName, category);
            });
            return filterId;
        }
        
        public void UpdateFilterName(string oldName, string category, string newName)
        {
            if (string.IsNullOrWhiteSpace(oldName) || string.IsNullOrWhiteSpace(newName) || string.IsNullOrWhiteSpace(category))
                return;
            
            UseFilterRepository(repo => repo.UpdateFilterName(oldName, category, newName));
        }
        
        public void DeleteFilter(string filterName, string category)
        {
            if (string.IsNullOrWhiteSpace(filterName) || string.IsNullOrWhiteSpace(category))
                return;
            
            UseFilterRepository(repo => repo.DeleteFilter(filterName, category));
        }
        
        public List<FilterInfo> GetAllFilters()
        {
            var filters = new List<FilterInfo>();
            UseFilterRepository(repo =>
            {
                var dbFilters = repo.GetAllFilters();
                foreach (var f in dbFilters)
                {
                    // Get FilterId for each filter
                    var filterId = repo.GetFilterId(f.FilterName, f.Category);
                    
                    filters.Add(new FilterInfo
                    {
                        FilterId = filterId,
                        FilterName = f.FilterName,
                        Category = f.Category
                    });
                }
            });
            return filters;
        }
        
        public int GetFilterId(string filterName, string category)
        {
            int filterId = -1;
            UseFilterRepository(repo =>
            {
                filterId = repo.GetFilterId(filterName, category);
            });
            return filterId;
        }
        
        public void SaveFilterUIState(string filterName, string category, List<string> selectedHostCategories, OpeningSettings settings)
        {
            UseFilterRepository(repo =>
            {
                repo.SaveFilterUIState(filterName, category, selectedHostCategories, settings);
            });
        }
        
        public (List<string> selectedHostCategories, OpeningSettings settings) LoadFilterUIState(string filterName, string category)
        {
            List<string> hostCategories = null;
            OpeningSettings settings = null;
            
            UseFilterRepository(repo =>
            {
                var result = repo.LoadFilterUIState(filterName, category);
                // FilterRepository returns 5-element tuple, but interface only needs 2
                hostCategories = result.SelectedHostCategories;
                settings = result.OpeningSettings;
            });
            
            return (hostCategories, settings);
        }
        
        private void UseFilterRepository(Action<FilterRepository> action)
        {
            if (action == null) return;
            
            try
            {
                using (var context = new SleeveDbContext(_document, msg => _logger.Info($"[FilterRepositoryAdapter] {msg}")))
                {
                    var repository = new FilterRepository(context, msg => _logger.Info($"[FilterRepositoryAdapter] {msg}"));
                    action(repository);
                }
            }
            catch (Exception ex)
            {
                _logger.Error($"FilterRepository operation failed: {ex.Message}", ex, "FilterRepositoryAdapter");
                throw;
            }
        }
    }
}
