using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Service for applying MEPMARK parameters to cluster sleeves
    /// Handles continuous numbering to avoid duplicates on re-runs
    /// </summary>
    public class MarkParameterService
    {
        // ⚠️ PERFORMANCE: Cache clash zones to avoid O(n·m) XML deserialization
        private Dictionary<long, ClashZone> _clashZoneCache;
        private bool _cacheInitialized = false;
        private Document _cachedDocument = null; // Track document for cache invalidation
        
        /// <summary>
        /// Apply MEPMARK to cluster sleeves for a specific category
        /// </summary>
        /// <param name="doc">Revit document</param>
        /// <param name="category">Target category (Ducts, Pipes, Cable Trays, etc.)</param>
        /// <param name="projectPrefix">Project prefix (e.g., "SLEEVE_")</param>
        /// <param name="disciplinePrefix">Discipline prefix (e.g., "DCT", "PLU", "ELE")</param>
        /// <returns>Tuple of (processedCount, errorCount)</returns>
        /// <summary>
        /// Ensure shared parameters are loaded into the project
        /// </summary>
        private void EnsureSharedParametersLoaded(Document doc)
        {
            try
            {
                // ✅ FIX: Get Resources path relative to DLL location (for deployment)
                var addinLocation = System.Reflection.Assembly.GetExecutingAssembly().Location;
                var addinDirectory = Path.GetDirectoryName(addinLocation);
                string sharedParamFile = Path.Combine(addinDirectory ?? "", "Resources", "Opening family shared parameter.txt");

                if (!File.Exists(sharedParamFile))
                {
                    DebugLogger.Warning($"[MarkParameterService] Shared parameter file not found: {sharedParamFile}");
                    return;
                }

                // Check if shared parameter file is already loaded
                var currentSharedParams = doc.Application.SharedParametersFilename;
                if (!string.IsNullOrEmpty(currentSharedParams) && currentSharedParams.Contains("Opening family shared parameter.txt"))
                {
                    DebugLogger.Info($"[MarkParameterService] Shared parameters already loaded: {currentSharedParams}");
                    return;
                }

                // Load the shared parameter file
                doc.Application.SharedParametersFilename = sharedParamFile;
                DebugLogger.Info($"[MarkParameterService] Loaded shared parameter file: {sharedParamFile}");

                // Create a group for the parameters if it doesn't exist
                var groupName = "Openings";
                var sharedParams = doc.Application.OpenSharedParameterFile();

                if (sharedParams != null)
                {
                    var group = sharedParams.Groups.get_Item(groupName);
                    if (group == null)
                    {
                        // Create the group if it doesn't exist
                        group = sharedParams.Groups.Create(groupName);
                        DebugLogger.Info($"[MarkParameterService] Created shared parameter group: {groupName}");
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[MarkParameterService] Error loading shared parameters: {ex.Message}");
                // Continue anyway - parameters might already be loaded
            }
        }

        public (int processedCount, int errorCount) ApplyMepMarkToClusters(
            Document doc, string category, string projectPrefix, string disciplinePrefix, bool remarkAll = false)
        {
            int processedCount = 0;
            int errorCount = 0;

            try
            {
                // ✅ FIX: Use SafeFileLogger for deployment-compatible log path
                string mepmarkLogPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                File.AppendAllText(mepmarkLogPath, $"\n===== MEPMARK DEBUG SESSION STARTED {DateTime.Now:yyyy-MM-dd HH:mm:ss} =====\n");
                File.AppendAllText(mepmarkLogPath, $"Category: {category}\n");
                File.AppendAllText(mepmarkLogPath, $"Project Prefix: {projectPrefix}\n");
                File.AppendAllText(mepmarkLogPath, $"Discipline Prefix: {disciplinePrefix}\n");
                
                // ⚠️ CRITICAL: Ensure shared parameters are loaded into the project
                EnsureSharedParametersLoaded(doc);

                // ✅ UPDATED: Get BOTH cluster sleeves AND individual sleeves for this category
                var clusterSleeves = GetClusterSleevesForCategory(doc, category);
                var individualSleeves = GetIndividualSleevesForCategory(doc, category);
                var allSleeves = new List<FamilyInstance>();
                allSleeves.AddRange(clusterSleeves);
                allSleeves.AddRange(individualSleeves);
                
                DebugLogger.Info($"[MarkParameterService] Found {clusterSleeves.Count} cluster sleeves + {individualSleeves.Count} individual sleeves = {allSleeves.Count} total for category '{category}'");
                File.AppendAllText(mepmarkLogPath, $"Found {clusterSleeves.Count} cluster sleeves for category '{category}'\n");
                File.AppendAllText(mepmarkLogPath, $"Found {individualSleeves.Count} individual sleeves for category '{category}'\n");
                File.AppendAllText(mepmarkLogPath, $"Total sleeves to mark: {allSleeves.Count}\n");
                
                if (allSleeves.Count == 0)
                {
                    DebugLogger.Warning($"[MarkParameterService] No sleeves found for category '{category}'");
                    File.AppendAllText(mepmarkLogPath, $"WARNING: No sleeves found for category '{category}'\n");
                    return (0, 0);
                }
                
                // Get starting index based on existing marks (continuous numbering)
                int startIndex = GetMaxExistingMarkNumber(doc, projectPrefix, disciplinePrefix) + 1;
                
                DebugLogger.Info($"[MarkParameterService] Starting mark numbering at: {startIndex} (based on existing marks)");
                File.AppendAllText(mepmarkLogPath, $"Starting mark numbering at: {startIndex} (based on existing marks)\n");
                
                // Apply MEPMARK to each sleeve (both cluster and individual)
                int actualIndex = startIndex;
                for (int i = 0; i < allSleeves.Count; i++)
                {
                    try
                    {
                        var sleeve = allSleeves[i];
                        
                        // ✅ FIX: Skip sleeves that already have MEPMARK value (unless RemarkAll is true)
                        var existingMark = sleeve.LookupParameter("MEP Mark")?.AsString() ?? 
                                          sleeve.LookupParameter("Mark")?.AsString();
                        
                        if (!remarkAll && !string.IsNullOrEmpty(existingMark))
                        {
                            File.AppendAllText(mepmarkLogPath, $"SKIP sleeve {sleeve.Id}: already has mark '{existingMark}' (RemarkAll=false)\n");
                            continue; // Skip - already marked
                        }
                        
                        if (remarkAll && !string.IsNullOrEmpty(existingMark))
                        {
                            File.AppendAllText(mepmarkLogPath, $"OVERWRITE sleeve {sleeve.Id}: changing '{existingMark}' → (RemarkAll=true)\n");
                        }
                        
                        // ✅ ENHANCED: Log detailed info about sleeve before marking
                        var mepElementIdParam = sleeve.LookupParameter("MEP_ElementId");
                        long mepElementId = mepElementIdParam?.AsInteger() ?? -1;
                        var clashZone = GetClashZoneByMepElementId(mepElementId, doc);
                        
                        File.AppendAllText(mepmarkLogPath, 
                            $"[MARK-ASSIGN] Sleeve {sleeve.Id}: MEP_ID={mepElementId}, Category={clashZone?.MepElementCategory ?? "UNKNOWN"}, " +
                            $"IsCluster={clashZone?.IsClusterResolved ?? false}, ClusterId={clashZone?.ClusterSleeveInstanceId ?? -1}\n");
                        
                        string markValue = GenerateMarkValue(disciplinePrefix, actualIndex);
                        string fullMarkValue = $"{projectPrefix}{markValue}";
                        
                        File.AppendAllText(mepmarkLogPath, 
                            $"[MARK-ASSIGN] Generating mark: prefix='{projectPrefix}', discipline='{disciplinePrefix}', index={actualIndex} → '{fullMarkValue}'\n");
                        
                        SetMarkParameter(sleeve, fullMarkValue);
                        processedCount++;
                        actualIndex++;
                        
                        DebugLogger.Info($"[MarkParameterService] Applied MEPMARK '{fullMarkValue}' to sleeve {sleeve.Id.IntegerValue}");
                        File.AppendAllText(mepmarkLogPath, $"[MARK-ASSIGN] ✅ SUCCESS: Applied MEPMARK '{fullMarkValue}' to sleeve {sleeve.Id.IntegerValue}\n");
                    }
                    catch (Exception ex)
                    {
                        errorCount++;
                        DebugLogger.Error($"[MarkParameterService] Error applying MEPMARK to sleeve {allSleeves[i].Id.IntegerValue}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[MarkParameterService] Error processing category '{category}': {ex.Message}");
                throw;
            }
            
            string mepmarkLogPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
            File.AppendAllText(mepmarkLogPath, $"MEPMARK session completed: {processedCount} processed, {errorCount} errors\n");
            File.AppendAllText(mepmarkLogPath, $"===== MEPMARK DEBUG SESSION ENDED {DateTime.Now:yyyy-MM-dd HH:mm:ss} =====\n\n");
            
            return (processedCount, errorCount);
        }
        
        /// <summary>
        /// Find individual sleeves for a specific category (non-clustered sleeves)
        /// </summary>
        private List<FamilyInstance> GetIndividualSleevesForCategory(Document doc, string category)
        {
            var allSleeves = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => {
                    var famName = fi.Symbol?.Family?.Name ?? string.Empty;
                    // ✅ Use contains instead of exact match
                    return famName.IndexOf("OpeningOnWall", StringComparison.OrdinalIgnoreCase) >= 0
                        || famName.IndexOf("OpeningOnSlab", StringComparison.OrdinalIgnoreCase) >= 0;
                })
                .ToList();

            // ✅ FIX: Find individual sleeves by checking what EXISTS in Revit and is NOT a cluster sleeve
            var individualSleeves = new List<FamilyInstance>();

            foreach (var sleeve in allSleeves)
            {
                var mepElementIdParam = sleeve.LookupParameter("MEP_ElementId");
                if (mepElementIdParam != null)
                {
                    long mepElementId = mepElementIdParam.AsInteger();
                    string mepmarkLogPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                    File.AppendAllText(mepmarkLogPath, 
                        $"[INDIVIDUAL-DEBUG] Checking sleeve {sleeve.Id}: MEP_ElementId={mepElementId}\n");
                    
                    // Find matching clash zone (pass doc context for project-specific paths)
                    var clashZone = GetClashZoneByMepElementId(mepElementId, doc);
                    File.AppendAllText(mepmarkLogPath, 
                        $"[INDIVIDUAL-DEBUG] Sleeve {sleeve.Id}: ClashZone={clashZone != null}, Category={clashZone?.MepElementCategory}\n");
                    
                    if (clashZone != null && clashZone.MepElementCategory == category)
                    {
                        // ✅ Individual sleeve = NOT part of a cluster (check ClusterSleeveInstanceId)
                        // If this sleeve's ID doesn't match any ClusterSleeveInstanceId, it's an individual sleeve
                        bool isClusterSleeve = (clashZone.IsClusterResolved && 
                                               clashZone.ClusterSleeveInstanceId > 0 && 
                                               sleeve.Id.IntegerValue == clashZone.ClusterSleeveInstanceId);
                        
                        File.AppendAllText(mepmarkLogPath, 
                            $"[INDIVIDUAL-DEBUG] Sleeve {sleeve.Id}: MEP_ID={mepElementId}, " +
                            $"Category_match={clashZone.MepElementCategory == category} (XML='{clashZone.MepElementCategory}' vs Requested='{category}'), " +
                            $"isClusterSleeve={isClusterSleeve} (IsClusterResolved={clashZone.IsClusterResolved}, ClusterSleeveInstanceId={clashZone.ClusterSleeveInstanceId}, SleeveId={sleeve.Id.IntegerValue})\n");
                        
                        if (!isClusterSleeve)
                        {
                            if (clashZone.MepElementCategory == category)
                            {
                                individualSleeves.Add(sleeve);
                                File.AppendAllText(mepmarkLogPath, 
                                    $"[INDIVIDUAL-DEBUG] ✅ ACCEPTED: Individual sleeve {sleeve.Id} for category '{category}'\n");
                            }
                            else
                            {
                                File.AppendAllText(mepmarkLogPath, 
                                    $"[INDIVIDUAL-DEBUG] ❌ REJECTED: Category mismatch - XML='{clashZone.MepElementCategory}' vs Requested='{category}'\n");
                            }
                        }
                        else
                        {
                            File.AppendAllText(mepmarkLogPath, 
                                $"[INDIVIDUAL-DEBUG] ❌ REJECTED: This is a cluster sleeve, not individual\n");
                        }
                    }
                }
            }

            string mepmarkLogPathFinal = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
            File.AppendAllText(mepmarkLogPathFinal, 
                $"Found {individualSleeves.Count} individual sleeves for category '{category}'\n");
            return individualSleeves;
        }

        /// <summary>
        /// Find cluster sleeves for a specific category
        /// Uses IsClusterResolved flag to identify actual cluster sleeves
        /// </summary>
        private List<FamilyInstance> GetClusterSleevesForCategory(Document doc, string category)
        {
            // ✅ DEBUG: Log all Opening families in model to verify family names
            var allOpeningFamilies = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol?.Family?.Name?.Contains("Opening") == true)
                .ToList();

            string mepmarkLogPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
            File.AppendAllText(mepmarkLogPath, 
                $"DEBUG: total Opening families in model = {allOpeningFamilies.Count}\n" +
                $"DEBUG: exact names found = {string.Join(", ", allOpeningFamilies.Select(f => f.Symbol.Family.Name).Distinct())}\n");

            var allClusterSleeves = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => {
                    var famName = fi.Symbol?.Family?.Name ?? string.Empty;
                    // ✅ FIX: Use contains instead of exact match to handle any family name variations
                    return famName.IndexOf("OpeningOnWall", StringComparison.OrdinalIgnoreCase) >= 0
                        || famName.IndexOf("OpeningOnSlab", StringComparison.OrdinalIgnoreCase) >= 0;
                })
                .ToList();

            File.AppendAllText(mepmarkLogPath, $"Found {allClusterSleeves.Count} total opening sleeves (after family name filter)\n");

            // ✅ NEW APPROACH: Find cluster sleeves using IsClusterResolved flag
            var categorySleeves = new List<FamilyInstance>();

            foreach (var sleeve in allClusterSleeves)
            {
                var mepElementIdParam = sleeve.LookupParameter("MEP_ElementId");
                if (mepElementIdParam != null)
                {
                    long mepElementId = mepElementIdParam.AsInteger();
                    
                    // Find matching clash zone in XML files (pass doc context for project-specific paths)
                    var clashZone = GetClashZoneByMepElementId(mepElementId, doc);
                    
                    string mepmarkLogPathDebug = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                    // ⚠️ DIAGNOSTIC: Log exactly why sleeves are rejected
                    if (clashZone == null)
                        File.AppendAllText(mepmarkLogPathDebug,
                            $"REJECT {sleeve.Id}: no clash-zone for MEPid {mepElementId}\n");
                    else if (!clashZone.IsClusterResolved)
                        File.AppendAllText(mepmarkLogPathDebug,
                            $"REJECT {sleeve.Id}: IsClusterResolved=false\n");
                    else if (clashZone.ClusterSleeveInstanceId != sleeve.Id.IntegerValue)
                        File.AppendAllText(mepmarkLogPathDebug,
                            $"REJECT {sleeve.Id}: ClusterSleeveInstanceId={clashZone.ClusterSleeveInstanceId} != {sleeve.Id.IntegerValue}\n");
                    else if (clashZone.MepElementCategory != category)
                        File.AppendAllText(mepmarkLogPathDebug,
                            $"REJECT {sleeve.Id}: category='{clashZone.MepElementCategory}' != '{category}'\n");
                    
                    if (clashZone != null && clashZone.IsClusterResolved && clashZone.ClusterSleeveInstanceId > 0)
                    {
                        // Check if this sleeve is the actual cluster sleeve
                        if (sleeve.Id.IntegerValue == clashZone.ClusterSleeveInstanceId)
                        {
                            // Verify category matches
                            bool categoryMatch = clashZone.MepElementCategory == category;
                            File.AppendAllText(mepmarkLogPathDebug, 
                                $"[CLUSTER-DEBUG] Category check: XML='{clashZone.MepElementCategory}' vs Requested='{category}' → Match={categoryMatch}\n");
                            
                            if (categoryMatch)
                            {
                                categorySleeves.Add(sleeve);
                                File.AppendAllText(mepmarkLogPathDebug, 
                                    $"[CLUSTER-DEBUG] ✅ ACCEPTED: Cluster sleeve {sleeve.Id} for category '{category}' " +
                                    $"(ClusterSleeveId={clashZone.ClusterSleeveInstanceId}, MEP_ID={mepElementId})\n");
                            }
                            else
                            {
                                File.AppendAllText(mepmarkLogPathDebug, 
                                    $"[CLUSTER-DEBUG] ❌ REJECTED: Category mismatch - XML='{clashZone.MepElementCategory}' vs Requested='{category}'\n");
                            }
                        }
                    }
                }
                else
                {
                    string mepmarkLogPathMissing = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                    File.AppendAllText(mepmarkLogPathMissing, $"Sleeve {sleeve.Id} missing MEP_ElementId parameter\n");
                }
            }

            string mepmarkLogPathFinal = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
            File.AppendAllText(mepmarkLogPathFinal, $"Found {categorySleeves.Count} cluster sleeves for category '{category}'\n");
            return categorySleeves;
        }

        /// <summary>
        /// Get clash zone by MEP element ID from XML files
        /// ⚠️ PERFORMANCE: Uses cache to avoid O(n·m) XML deserialization
        /// </summary>
        private ClashZone GetClashZoneByMepElementId(long mepElementId, Document doc = null)
        {
            // Initialize cache once per service instance, or reinitialize if document changed
            if (!_cacheInitialized || (doc != null && _cachedDocument != doc))
            {
                InitializeClashZoneCache(doc);
                _cacheInitialized = true;
                _cachedDocument = doc;
            }

            // Fast O(1) lookup
            if (_clashZoneCache != null && _clashZoneCache.TryGetValue(mepElementId, out ClashZone clashZone))
            {
                string mepmarkLogPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                File.AppendAllText(mepmarkLogPath, 
                    $"[CACHE-LOOKUP] ✓ Found clash zone for MEP_ID={mepElementId}: Category={clashZone.MepElementCategory}, " +
                    $"IsClusterResolved={clashZone.IsClusterResolved}, ClusterSleeveId={clashZone.ClusterSleeveInstanceId}\n");
                return clashZone;
            }

            string mepmarkLogPathNotFound = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
            File.AppendAllText(mepmarkLogPathNotFound, $"[CACHE-LOOKUP] ❌ NOT FOUND: MEP_ID={mepElementId} (cache size: {_clashZoneCache?.Count ?? 0})\n");
            return null;
        }

        /// <summary>
        /// Initialize clash zone cache from all XML files (called once)
        /// </summary>
        private void InitializeClashZoneCache(Document doc = null)
        {
            _clashZoneCache = new Dictionary<long, ClashZone>();

            try
            {
                // ✅ FIX: Use ProjectPathService to get project-specific filters directory
                string filtersDirectory;
                if (doc != null)
                {
                    filtersDirectory = ProjectPathService.GetFiltersDirectory(doc);
                }
                else
                {
                    // Fallback to Default project path if no document context
                    filtersDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings", "Projects", "Default", "Filters");
                }
                
                string mepmarkLogPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                File.AppendAllText(mepmarkLogPath, $"\n[CACHE-INIT] ===== CACHE INITIALIZATION STARTED =====\n");
                File.AppendAllText(mepmarkLogPath, $"[CACHE-INIT] Document: {(doc != null ? doc.Title : "NULL (using Default path)")}\n");
                File.AppendAllText(mepmarkLogPath, $"[CACHE-INIT] Filters Directory: {filtersDirectory}\n");
                
                if (!Directory.Exists(filtersDirectory))
                {
                    File.AppendAllText(mepmarkLogPath, $"[CACHE-INIT] ❌ ERROR: Filters directory not found: {filtersDirectory}\n");
                    return;
                }

                var allXmlFiles = Directory.GetFiles(filtersDirectory, "*.xml");
                // ✅ CRITICAL FIX: Only load filtername_categoryname.xml files (OpeningFilter structure)
                // Exclude: *_global.xml (GlobalFlagStorage - flag management only)
                // Exclude: *_CONDITIONS.xml (Conditions - clearance settings only)
                var xmlFiles = allXmlFiles.Where(f => 
                {
                    string fileName = Path.GetFileName(f);
                    return !fileName.EndsWith("_global.xml", StringComparison.OrdinalIgnoreCase) &&
                           !fileName.EndsWith("_CONDITIONS.xml", StringComparison.OrdinalIgnoreCase) &&
                           !fileName.EndsWith("_conditions.xml", StringComparison.OrdinalIgnoreCase);
                }).ToList();
                
                File.AppendAllText(mepmarkLogPath, $"[CACHE-INIT] Found {allXmlFiles.Length} XML files total, {xmlFiles.Count} filter files (excluded {allXmlFiles.Length - xmlFiles.Count} global files)\n");
                
                // ✅ ENHANCED: List all XML files found
                foreach (var xmlFile in xmlFiles)
                {
                    File.AppendAllText(mepmarkLogPath, $"[CACHE-INIT]   - {Path.GetFileName(xmlFile)}\n");
                }

                int totalClashZones = 0;
                int filesProcessed = 0;
                foreach (var xmlFile in xmlFiles)
                {
                    try
                    {
                        int zonesInFile = 0;
                        var serializer = new System.Xml.Serialization.XmlSerializer(typeof(OpeningFilter));
                        using (var reader = new StreamReader(xmlFile))
                        {
                            var filter = (OpeningFilter)serializer.Deserialize(reader);
                            if (filter?.ClashZoneStorage?.ClashZones != null)
                            {
                                File.AppendAllText(mepmarkLogPath, $"[CACHE-INIT] Loading {Path.GetFileName(xmlFile)}: {filter.ClashZoneStorage.ClashZones.Count} clash zones\n");
                                
                                foreach (var clashZone in filter.ClashZoneStorage.ClashZones)
                                {
                                    long key = clashZone.MepElementId.IntegerValue;
                                    if (!_clashZoneCache.ContainsKey(key))
                                    {
                                        _clashZoneCache[key] = clashZone;
                                        totalClashZones++;
                                        zonesInFile++;
                                        
                                        // ✅ ENHANCED: Log key clash zone details for first few zones
                                        if (zonesInFile <= 3)
                                        {
                                            File.AppendAllText(mepmarkLogPath, 
                                                $"[CACHE-INIT]   Zone {zonesInFile}: MEP_ID={key}, Category={clashZone.MepElementCategory}, " +
                                                $"IsClusterResolved={clashZone.IsClusterResolved}, ClusterSleeveId={clashZone.ClusterSleeveInstanceId}, " +
                                                $"SleeveId={clashZone.SleeveInstanceId}\n");
                                        }
                                    }
                                }
                                
                                File.AppendAllText(mepmarkLogPath, $"[CACHE-INIT] ✓ Loaded {zonesInFile} zones from {Path.GetFileName(xmlFile)}\n");
                                filesProcessed++;
                            }
                            else
                            {
                                File.AppendAllText(mepmarkLogPath, $"[CACHE-INIT] ⚠ {Path.GetFileName(xmlFile)}: No ClashZoneStorage or ClashZones found\n");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        File.AppendAllText(mepmarkLogPath, $"[CACHE-INIT] ❌ ERROR loading {Path.GetFileName(xmlFile)}: {ex.Message}\n");
                        File.AppendAllText(mepmarkLogPath, $"[CACHE-INIT]   Stack: {ex.StackTrace}\n");
                        continue;
                    }
                }

                File.AppendAllText(mepmarkLogPath, $"[CACHE-INIT] ===== CACHE INITIALIZATION COMPLETE =====\n");
                File.AppendAllText(mepmarkLogPath, $"[CACHE-INIT] ✓ Cached {totalClashZones} unique clash zones from {filesProcessed}/{xmlFiles.Length} files\n");
                File.AppendAllText(mepmarkLogPath, $"[CACHE-INIT] Cache size: {_clashZoneCache.Count} entries\n\n");
            }
            catch (Exception ex)
            {
                string mepmarkLogPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                File.AppendAllText(mepmarkLogPath, $"[CACHE-INIT] ❌ FATAL ERROR initializing cache: {ex.Message}\n");
                File.AppendAllText(mepmarkLogPath, $"[CACHE-INIT] Stack: {ex.StackTrace}\n");
                DebugLogger.Error($"[MarkParameterService] Error initializing clash zone cache: {ex.Message}");
            }
        }

        /// <summary>
        /// Get category from MEP Element ID by searching XML files
        /// </summary>
        private string GetCategoryFromMepElementId(long mepElementId)
        {
            try
            {
                var filtersDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings", "Projects", "Default", "Filters");

                // Search all XML files for this MEP element ID
                var xmlFiles = Directory.GetFiles(filtersDirectory, "*.xml");

                foreach (var xmlFile in xmlFiles)
                {
                    try
                    {
                        var serializer = new System.Xml.Serialization.XmlSerializer(typeof(OpeningFilter));
                        using (var reader = new StreamReader(xmlFile))
                        {
                            var filter = (OpeningFilter)serializer.Deserialize(reader);
                            if (filter?.ClashZoneStorage?.ClashZones != null)
                            {
                                foreach (var clashZone in filter.ClashZoneStorage.ClashZones)
                                {
                                    if (clashZone.MepElementId.IntegerValue == mepElementId)
                                    {
                                        return clashZone.MepElementCategory;
                                    }
                                }
                            }
                        }
                    }
                    catch
                    {
                        // Ignore deserialization errors for individual files
                        continue;
                    }
                }

                    DebugLogger.Warning($"[MarkParameterService] No category found for MEP Element ID {mepElementId} in {xmlFiles.Length} XML files");
                    string mepmarkLogPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                    File.AppendAllText(mepmarkLogPath, $"WARNING: No category found for MEP Element ID {mepElementId} in {xmlFiles.Length} XML files\n");
                    return "Unknown";
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[MarkParameterService] Error getting category for MEP Element ID {mepElementId}: {ex.Message}");
                return "Unknown";
            }
        }
        
        /// <summary>
        /// ✅ OPTIMIZED: Get maximum existing mark number to continue numbering
        /// Only filters sleeve family instances for better performance
        /// </summary>
        private int GetMaxExistingMarkNumber(Document doc, string projectPrefix, string disciplinePrefix)
        {
            try
            {
                string expectedPrefix = $"{projectPrefix}{disciplinePrefix}";
                int maxNumber = 0;
                
                // ✅ OPTIMIZED: Only get sleeve family instances (performance optimization)
                var sleeveElements = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(fi => {
                        var famName = fi.Symbol?.Family?.Name ?? string.Empty;
                        return famName == "RectangularOpeningOnWall" || 
                               famName == "CircularOpeningOnWall" ||
                               famName == "RectangularOpeningOnSlab" || 
                               famName == "CircularOpeningOnSlab";
                    })
                    .ToList();
                
                foreach (var element in sleeveElements)
                {
                    // Resolve mark parameter case/space/underscore-insensitively
                    var markParam = ResolveMarkParameter(element);
                    if (markParam != null)
                    {
                        string markValue = markParam.AsString() ?? "";

                        // Check if this mark matches our pattern: ProjectPrefix + DisciplinePrefix + Number
                        if (markValue.StartsWith(expectedPrefix, StringComparison.OrdinalIgnoreCase))
                        {
                            // Extract the number part
                            string numberPart = markValue.Substring(expectedPrefix.Length);
                            if (int.TryParse(numberPart, out int number))
                            {
                                maxNumber = Math.Max(maxNumber, number);
                            }
                        }
                    }
                }
                
                DebugLogger.Info($"[MarkParameterService] Max existing mark number for '{expectedPrefix}': {maxNumber}");
                return maxNumber;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[MarkParameterService] Error getting max existing mark number: {ex.Message}");
                return 0; // Start from 1 if error
            }
        }
        
        /// <summary>
        /// Generate mark value based on discipline prefix and number
        /// </summary>
        private string GenerateMarkValue(string disciplinePrefix, int number)
        {
            return $"{disciplinePrefix}{number:000}";
        }
        
        /// <summary>
        /// Set MEP Mark parameter on element
        /// </summary>
        private void SetMarkParameter(Element element, string markValue)
        {
            // Resolve parameter name case/space/underscore-insensitively
            var markParam = ResolveMarkParameter(element);
            string mepmarkLogPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");

            File.AppendAllText(mepmarkLogPath, 
                $"[SET-DEBUG] Element {element.Id}: ResolvedParam='{markParam?.Definition?.Name}', Value='{markValue}', Doc.IsModifiable={element.Document.IsModifiable}\n");
            
            if (markParam != null && !markParam.IsReadOnly && markParam.StorageType == StorageType.String)
            {
                markParam.Set(markValue);
                // ✅ FIX: Remove Regenerate() calls - not needed and can cause transaction issues
                // Parameter.Set() is immediate within a transaction, no need to regenerate
                var readBack = markParam.AsString();
                
                // Verify the value was set correctly
                if (string.Equals(readBack, markValue, StringComparison.Ordinal))
                {
                    File.AppendAllText(mepmarkLogPath, 
                        $"[SET-SUCCESS] Applied '{markValue}' to element {element.Id} using '{markParam.Definition?.Name}'\n");
                }
                else
                {
                    File.AppendAllText(mepmarkLogPath, 
                        $"[SET-VERIFY-FAIL] Element {element.Id}: attempted '{markValue}', read-back='{readBack ?? "<null>"}'\n");
                    throw new InvalidOperationException($"MEP Mark write did not persist on element {element.Id.IntegerValue}");
                }
            }
            else
            {
                var readonlyInfo = markParam != null ? $"Param='{markParam.Definition?.Name}', Readonly={markParam.IsReadOnly}, Type={markParam.StorageType}" : "Param=null";
                File.AppendAllText(mepmarkLogPath, 
                    $"[SET-FAIL] Element {element.Id}: {readonlyInfo}\n");
                throw new InvalidOperationException($"Cannot set MEP Mark parameter on element {element.Id.IntegerValue}: parameter missing or not writable");
            }
        }

        private Parameter ResolveMarkParameter(Element element)
        {
            // Accept: "MEP Mark", "MEP_Mark", "mep mark", "mep_mark" (case-insensitive), or fallback "Mark"
            var preferredNames = new[] { "mepmark", "mep_mark", "mep mark" };
            var fallbackNames = new[] { "mark" };

            try
            {
                // Search all instance parameters case/space/underscore-insensitively
                foreach (Parameter p in element.Parameters)
                {
                    var defName = p.Definition?.Name ?? string.Empty;
                    var norm = NormalizeName(defName);
                    if (preferredNames.Any(n => NormalizeName(n) == norm))
                    {
                        return p;
                    }
                }
                // Fallback to simple "Mark"
                foreach (Parameter p in element.Parameters)
                {
                    var defName = p.Definition?.Name ?? string.Empty;
                    var norm = NormalizeName(defName);
                    if (fallbackNames.Any(n => NormalizeName(n) == norm))
                    {
                        return p;
                    }
                }
            }
            catch { }
            
            // Final fallback: direct lookups (in case)
            return element.LookupParameter("MEP Mark") ?? element.LookupParameter("MEP_Mark") ?? element.LookupParameter("Mark");
        }

        private string NormalizeName(string s)
        {
            return (s ?? string.Empty).ToLowerInvariant().Replace(" ", string.Empty).Replace("_", string.Empty).Trim();
        }
    }
}

