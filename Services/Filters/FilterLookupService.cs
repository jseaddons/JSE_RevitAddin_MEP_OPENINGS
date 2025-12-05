using System;
using System.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Utils;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Filters
{
    /// <summary>
    /// ✅ SOLID COMPLIANCE (SRP): Service responsible for looking up filter information from database.
    /// Encapsulates all filter-related database queries.
    /// </summary>
    public class FilterLookupService
    {
        private readonly SleeveDbContext _dbContext;

        public FilterLookupService(SleeveDbContext dbContext)
        {
            _dbContext = dbContext ?? throw new ArgumentNullException(nameof(dbContext));
        }

        /// <summary>
        /// ✅ SRP: Gets FilterId from database by filter name and category.
        /// Includes performance monitoring, safety features, and comprehensive error handling.
        /// </summary>
        /// <param name="filterName">Filter name (will be normalized)</param>
        /// <param name="category">MEP category (optional)</param>
        /// <returns>FilterId if found, -1 otherwise</returns>
        public int GetFilterId(string filterName, string category = null)
        {
            // ✅ PERFORMANCE MONITORING: Track lookup time
            var timer = System.Diagnostics.Stopwatch.StartNew();
            
            try
            {
                // ✅ SAFETY: Input validation
                if (string.IsNullOrWhiteSpace(filterName))
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning("[FilterLookupService] Filter name is null or empty");
                    return -1;
                }

                // ✅ SAFETY: Validate database connection
                if (_dbContext?.Connection == null)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Error("[FilterLookupService] Database connection is null");
                    return -1;
                }

                // ✅ Normalize filter name (removes .xml and category suffixes)
                string filterNameForLookup = FilterNameHelper.NormalizeBaseName(filterName);

                using (var filterCmd = _dbContext.Connection.CreateCommand())
                {
                    if (!string.IsNullOrWhiteSpace(category))
                    {
                        filterCmd.CommandText = @"
                            SELECT FilterId FROM Filters 
                            WHERE FilterName = @FilterName AND Category = @Category
                            LIMIT 1";
                        filterCmd.Parameters.AddWithValue("@FilterName", filterNameForLookup);
                        filterCmd.Parameters.AddWithValue("@Category", category);
                    }
                    else
                    {
                        filterCmd.CommandText = @"
                            SELECT FilterId FROM Filters 
                            WHERE FilterName = @FilterName 
                            LIMIT 1";
                        filterCmd.Parameters.AddWithValue("@FilterName", filterNameForLookup);
                    }

                    var filterResult = filterCmd.ExecuteScalar();
                    timer.Stop();
                    
                    if (filterResult != null && filterResult != DBNull.Value)
                    {
                        int filterId = Convert.ToInt32(filterResult);
                        
                        // ✅ PERFORMANCE MONITORING: Log lookup time
                        if (!DeploymentConfiguration.DeploymentMode && timer.ElapsedMilliseconds > 50)
                        {
                            DebugLogger.Info($"[FilterLookupService] Filter lookup took {timer.ElapsedMilliseconds}ms: '{filterName}' -> FilterId={filterId}");
                        }
                        
                        return filterId;
                    }
                    else
                    {
                        timer.Stop();
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Warning($"[FilterLookupService] FilterId not found for '{filterName}' (category='{category}') in {timer.ElapsedMilliseconds}ms");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                timer.Stop();
                // ✅ SAFETY: Comprehensive error handling
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[FilterLookupService] Error looking up FilterId for '{filterName}': {ex.Message}");
                    DebugLogger.Error($"[FilterLookupService] Stack trace: {ex.StackTrace}");
                }
            }

            return -1;
        }
    }
}

