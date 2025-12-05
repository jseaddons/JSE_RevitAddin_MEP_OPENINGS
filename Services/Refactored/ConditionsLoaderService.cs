using System;
using System.IO;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Utils;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Refactored
{
    /// <summary>
    /// ✅ SOLID REFACTORED: Service for loading and saving opening conditions from XML files.
    /// Extracted from UniversalSleevePlacementCommand to adhere to SRP.
    /// </summary>
    public class ConditionsLoaderService : IConditionsLoader
    {
        private readonly Document _doc;
        private readonly ConditionsService _conditionsService;

        public ConditionsLoaderService(Document doc)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            var projectFiltersDir = ProjectPathService.GetFiltersDirectory(_doc);
            _conditionsService = new ConditionsService(_doc, projectFiltersDir, msg => 
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info(msg);
            });
        }

        public OpeningConditions LoadConditions(string filterName, string category)
        {
            try
            {
                // ✅ STANDARDIZED: Use normalized category names to match clash zone file naming
                string normalizedCategory = NormalizeCategoryName(category);
                string combinedKey = $"{filterName}_{normalizedCategory}";

                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[ConditionsLoader] Using combined key '{combinedKey}' (Filter: '{filterName}', Category: '{category}')");
                }

                var conditions = _conditionsService.LoadConditions(combinedKey);

                // Ensure CONDITIONS.xml exists in the project Filters directory; create if missing
                try
                {
                    var expectedPath = _conditionsService.GetConditionsFilePath(combinedKey);
                    if (!File.Exists(expectedPath))
                    {
                        // Initialize sane defaults tied to this filter/category key
                        if (conditions == null)
                            conditions = new OpeningConditions();
                        conditions.FilterName = filterName;
                        conditions.Category = category;
                        // Save immediately so subsequent runs find it
                        _conditionsService.SaveConditions(conditions, combinedKey);
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[ConditionsLoader] CONDITIONS.xml created at: {expectedPath}");
                        }
                    }
                }
                catch { }

                if (conditions != null)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[ConditionsLoader] Loaded conditions for '{combinedKey}'");
                    }
                }
                else
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Warning($"[ConditionsLoader] No conditions found for '{combinedKey}' - using defaults");
                    }
                    conditions = new OpeningConditions { FilterName = filterName, Category = category };
                }

                return conditions;
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[ConditionsLoader] Error loading conditions: {ex.Message}");
                }
                return new OpeningConditions { FilterName = filterName, Category = category };
            }
        }

        public string GetConditionsFilePath(string key)
        {
            return _conditionsService.GetConditionsFilePath(key);
        }

        public void SaveConditions(OpeningConditions conditions, string key)
        {
            _conditionsService.SaveConditions(conditions, key);
        }

        private string NormalizeCategoryName(string categoryName)
        {
            if (string.IsNullOrEmpty(categoryName))
                return "unknown";

            return categoryName
                .Replace(" ", "_")
                .ToLowerInvariant();
        }
    }
}

