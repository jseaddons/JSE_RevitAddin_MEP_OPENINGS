using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Service for applying MEPMARK parameters to cluster sleeves
    /// Handles continuous numbering to avoid duplicates on re-runs
    /// </summary>
    public class MarkParameterService
    {
        // ⚠️ PERFORMANCE: Cache clash zones to avoid O(n·m) XML deserialization
        private Dictionary<long, ClashZone>? _clashZoneCache;
        private bool _cacheInitialized = false;
        private Document? _cachedDocument = null; // Track document for cache invalidation
        
        // ✅ OPTIMIZATION: Cache max mark numbers per category+prefix combination to avoid repeated database queries
        private Dictionary<string, int>? _maxMarkNumberCache;
        private Document? _maxMarkCacheDocument = null;
        
        /// <summary>
        /// Apply MEPMARK to cluster sleeves for a specific category
        /// </summary>
        /// <param name="doc">Revit document</param>
        /// <param name="category">Target category (Ducts, Pipes, Cable Trays, etc.)</param>
        /// <param name="projectPrefix">Project prefix (e.g., "SLEEVE_")</param>
        /// <param name="disciplinePrefix">Discipline prefix (e.g., "DCT", "PLU", "ELE")</param>
        /// <param name="remarkAll">If true, re-apply marks even if they already exist</param>
        /// <param name="numberFormat">Number format (e.g., "000" for 001, 002, etc.)</param>
        /// <param name="markPrefixes">Optional MarkPrefixSettings to check RemarkProjectPrefix flag</param>
        /// <returns>Tuple of (processedCount, errorCount)</returns>
        public (int processedCount, int errorCount) ApplyMepMarkToClusters(
            Document doc, string category, string projectPrefix, string disciplinePrefix, bool remarkAll = false, string numberFormat = "000", MarkPrefixSettings? markPrefixes = null)
        {
            // Call the internal implementation
            return ApplyMepMarkToClustersInternal(doc, category, projectPrefix, disciplinePrefix, remarkAll, numberFormat, markPrefixes);
        }
        
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
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[MarkParameterService] Shared parameter file not found: {sharedParamFile}");
                    return;
                }

                // Check if shared parameter file is already loaded
                var currentSharedParams = doc.Application.SharedParametersFilename;
                if (!string.IsNullOrEmpty(currentSharedParams) && currentSharedParams.Contains("Opening family shared parameter.txt"))
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[MarkParameterService] Shared parameters already loaded: {currentSharedParams}");
                    return;
                }

                // Load the shared parameter file
                doc.Application.SharedParametersFilename = sharedParamFile;
                                if (!DeploymentConfiguration.DeploymentMode)
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
                                                if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[MarkParameterService] Created shared parameter group: {groupName}");
                    }
                }
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[MarkParameterService] Error loading shared parameters: {ex.Message}");
                // Continue anyway - parameters might already be loaded
            }
        }
        
        public (int processedCount, int errorCount) ApplyMepMarkToClustersInternal(
            Document doc, string category, string projectPrefix, string disciplinePrefix, bool remarkAll = false, string numberFormat = "000", MarkPrefixSettings? markPrefixes = null)
        {
            int processedCount = 0;
            int errorCount = 0;

            try
            {
                // ✅ FIX: Use SafeFileLogger for deployment-compatible log path
                string mepmarkLogPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                
                // ✅ BUILD TIME STAMP: Log to confirm latest DLL is loaded
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                var buildTime = System.IO.File.GetLastWriteTime(assembly.Location);
                                // ✅ DEPLOYMENT MODE: Skip file writes
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    File.AppendAllText(mepmarkLogPath, $"\n===== MEPMARK DEBUG SESSION STARTED {DateTime.Now:yyyy-MM-dd HH:mm:ss} =====\n");
                }
                File.AppendAllText(mepmarkLogPath, $"🔨 BUILD TIME: {buildTime:yyyy-MM-dd HH:mm:ss} (DLL: {Path.GetFileName(assembly.Location)})\n");
                                // ✅ DEPLOYMENT MODE: Skip file writes
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    File.AppendAllText(mepmarkLogPath, $"Category: {category}\n");
                }
                                // ✅ DEPLOYMENT MODE: Skip file writes
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    File.AppendAllText(mepmarkLogPath, $"Project Prefix: {projectPrefix}\n");
                }
                                // ✅ DEPLOYMENT MODE: Skip file writes
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    File.AppendAllText(mepmarkLogPath, $"Discipline Prefix: {disciplinePrefix}\n");
                }
                
                // ⚠️ CRITICAL: Ensure shared parameters are loaded into the project
                EnsureSharedParametersLoaded(doc);

                // ✅ UPDATED: Get BOTH cluster sleeves AND individual sleeves for this category
                var clusterSleeves = GetClusterSleevesForCategory(doc, category);
                var individualSleeves = GetIndividualSleevesForCategory(doc, category);
                var allSleeves = new List<FamilyInstance>();
                allSleeves.AddRange(clusterSleeves);
                allSleeves.AddRange(individualSleeves);
                
                // ✅ DEBUG: Log sleeve distribution by host type (Wall vs Floor)
                var sleevesByHostType = allSleeves.GroupBy(s => {
                    var famName = s.Symbol?.Family?.Name ?? "";
                    if (famName.IndexOf("OpeningOnWall", StringComparison.OrdinalIgnoreCase) >= 0) return "Wall";
                    if (famName.IndexOf("OpeningOnSlab", StringComparison.OrdinalIgnoreCase) >= 0) return "Floor";
                    return "Unknown";
                }).ToDictionary(g => g.Key, g => g.Count());
                
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[MarkParameterService] Found {clusterSleeves.Count} cluster sleeves + {individualSleeves.Count} individual sleeves = {allSleeves.Count} total for category '{category}'");
                                // ✅ DEPLOYMENT MODE: Skip file writes
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    File.AppendAllText(mepmarkLogPath, $"Found {clusterSleeves.Count} cluster sleeves for category '{category}'\n");
                }
                                // ✅ DEPLOYMENT MODE: Skip file writes
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    File.AppendAllText(mepmarkLogPath, $"Found {individualSleeves.Count} individual sleeves for category '{category}'\n");
                }
                                // ✅ DEPLOYMENT MODE: Skip file writes
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    File.AppendAllText(mepmarkLogPath, $"Total sleeves to mark: {allSleeves.Count}\n");
                    File.AppendAllText(mepmarkLogPath, $"Sleeves by host type: {string.Join(", ", sleevesByHostType.Select(kvp => $"{kvp.Key}={kvp.Value}"))}\n");
                    File.AppendAllText(mepmarkLogPath, $"RemarkAll flag: {remarkAll}\n");
                }
                
                if (allSleeves.Count == 0)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[MarkParameterService] No sleeves found for category '{category}'");
                                        // ✅ DEPLOYMENT MODE: Skip file writes
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        File.AppendAllText(mepmarkLogPath, $"WARNING: No sleeves found for category '{category}'\n");
                    }
                    return (0, 0);
                }
                
                // ✅ FIX: Get max number per category (consider all active prefixes for this category)
                var candidatePrefixes = GetCandidateDisciplinePrefixes(category, disciplinePrefix, markPrefixes);
                int categoryMaxNumber = GetMaxExistingMarkNumberForCategory(doc, category, candidatePrefixes);
                int startIndex = categoryMaxNumber + 1;
                
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[MarkParameterService] Category '{category}': Max existing number = {categoryMaxNumber}, Starting at: {startIndex}");
                                // ✅ DEPLOYMENT MODE: Skip file writes
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    File.AppendAllText(mepmarkLogPath, $"Category '{category}': Max existing number = {categoryMaxNumber}, Starting at: {startIndex}\n");
                }
                
                // ✅ TRACK: Keep track of used numbers when remark=true (to avoid duplicates)
                var usedNumbers = new HashSet<int>();
                if (remarkAll && categoryMaxNumber > 0)
                {
                    // Pre-populate with existing numbers
                    for (int n = 1; n <= categoryMaxNumber; n++)
                    {
                        usedNumbers.Add(n);
                    }
                }
                
                // Apply MEPMARK to each sleeve (both cluster and individual)
                int actualIndex = startIndex;
                for (int i = 0; i < allSleeves.Count; i++)
                {
                    try
                    {
                        var sleeve = allSleeves[i];
                        
                        // ✅ CRITICAL FIX: Get clash zone - try ClusterInstanceId first (for cluster sleeves), then MEP_ElementId (for individual sleeves)
                        ClashZone clashZone = null;
                        
                        // ✅ Declare variables once at the start of the loop iteration
                        var mepElementIdParam = sleeve.LookupParameter("MEP_ElementId");
                        long mepElementId = mepElementIdParam?.AsInteger() ?? -1;
                        
                        // First, try to get by ClusterInstanceId (for cluster sleeves)
                        int sleeveId = sleeve.Id.IntegerValue;
                        clashZone = GetClashZoneByClusterInstanceId(sleeveId, category, doc);
                        
                        // If not found, try MEP_ElementId (for individual sleeves)
                        if (clashZone == null && mepElementId > 0)
                        {
                            clashZone = GetClashZoneByMepElementId(mepElementId, doc);
                        }
                        
                        // ✅ ENHANCED: Resolve element-specific prefix (e.g., System Type override)
                        var elementPrefix = ResolveDisciplinePrefixForElement(category, disciplinePrefix, markPrefixes, clashZone);
                        if (!string.IsNullOrWhiteSpace(elementPrefix))
                        {
                            candidatePrefixes.Add(elementPrefix);
                        }

                        var existingMark = sleeve.LookupParameter("MEP Mark")?.AsString() ?? 
                                          sleeve.LookupParameter("Mark")?.AsString();
                        
                        int numberToUse;
                        
                        // ✅ FIX: When remark=true, extract number from existing mark (preserve number, update prefix only)
                        if (remarkAll && !string.IsNullOrEmpty(existingMark))
                        {
                            // Try to extract number from existing mark
                            int? extractedNumber = ExtractNumberFromMark(existingMark, candidatePrefixes);
                            if (extractedNumber.HasValue)
                            {
                                numberToUse = extractedNumber.Value;
                                usedNumbers.Add(numberToUse); // Track this number as used
                                                                // ✅ DEPLOYMENT MODE: Skip file writes
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    File.AppendAllText(mepmarkLogPath, 
                                    $"[MARK-ASSIGN] Remark=true: Sleeve {sleeve.Id} existing mark '{existingMark}' → extracted number {numberToUse}, updating prefix only\n");
                                }
                                // Don't increment actualIndex - preserving existing number
                            }
                            else
                            {
                                // Can't extract number, use next available number that's not used
                                while (usedNumbers.Contains(actualIndex))
                                {
                                    actualIndex++;
                                }
                                numberToUse = actualIndex;
                                usedNumbers.Add(numberToUse);
                                actualIndex++; // Increment for next sleeve
                                                                // ✅ DEPLOYMENT MODE: Skip file writes
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    File.AppendAllText(mepmarkLogPath, 
                                    $"[MARK-ASSIGN] Remark=true: Sleeve {sleeve.Id} existing mark '{existingMark}' → couldn't extract number, using next available: {numberToUse}\n");
                                }
                            }
                        }
                        else if (!remarkAll && !string.IsNullOrEmpty(existingMark))
                        {
                            // Skip - already marked and remark=false
                            File.AppendAllText(mepmarkLogPath, $"SKIP sleeve {sleeve.Id}: already has mark '{existingMark}' (remark=false)\n");
                            continue; // Skip - already marked
                        }
                        else
                        {
                            // No existing mark, use next number and increment
                            if (remarkAll)
                            {
                                // When remark=true, skip used numbers
                                while (usedNumbers.Contains(actualIndex))
                        {
                                    actualIndex++;
                                }
                                numberToUse = actualIndex;
                                usedNumbers.Add(numberToUse);
                            }
                            else
                            {
                                numberToUse = actualIndex;
                            }
                            actualIndex++; // Increment for next sleeve
                        }
                        
                        // ✅ ENHANCED: Log detailed info about sleeve before marking
                        
                        // ✅ DEBUG: Log sleeve host type
                        var famName = sleeve.Symbol?.Family?.Name ?? "";
                        var hostType = famName.IndexOf("OpeningOnWall", StringComparison.OrdinalIgnoreCase) >= 0 ? "Wall" :
                                      famName.IndexOf("OpeningOnSlab", StringComparison.OrdinalIgnoreCase) >= 0 ? "Floor" : "Unknown";
                        
                        // ✅ ENHANCED LOGGING: Log ALL sleeves (not just first 10) to diagnose floor sleeve issue
                        // Note: mepElementIdParam and mepElementId are already declared above
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            File.AppendAllText(mepmarkLogPath, 
                                $"[MARK-PROCESS] Sleeve {sleeve.Id}: HostType={hostType}, Family={famName}, MEP_ID={mepElementId}, " +
                                $"Category={clashZone?.MepElementCategory ?? "null"}, RequestedCategory={category}, " +
                                $"CategoryMatch={clashZone?.MepElementCategory == category}, RemarkAll={remarkAll}, ExistingMark='{existingMark ?? "null"}'\n");
                        }
                        
                        // ✅ CRITICAL FIX: Check if clash zone category matches requested category
                        // This ensures floor sleeves are only processed if their category matches
                        if (clashZone != null && clashZone.MepElementCategory != category)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                File.AppendAllText(mepmarkLogPath, 
                                    $"[MARK-PROCESS] ⚠️ SKIP: Sleeve {sleeve.Id} category mismatch - XML='{clashZone.MepElementCategory}' vs Requested='{category}'\n");
                            }
                            continue; // Skip sleeves that don't match the requested category
                        }
                        
                                                // ✅ DEPLOYMENT MODE: Skip file writes
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            File.AppendAllText(mepmarkLogPath, 
                            $"[MARK-ASSIGN] Sleeve {sleeve.Id}: MEP_ID={mepElementId}, Category={clashZone?.MepElementCategory ?? "UNKNOWN"}, " +
                            $"IsCluster={clashZone?.IsClusterResolved ?? false}, ClusterId={clashZone?.ClusterSleeveInstanceId ?? -1}\n");
                        }
                        
                        string markValue = GenerateMarkValue(elementPrefix, numberToUse, numberFormat);
                        
                        // ✅ CRITICAL FIX: Preserve project prefix unless user explicitly opted to remark it
                        string effectiveProjectPrefix = projectPrefix;
                        if (markPrefixes != null && !markPrefixes.RemarkProjectPrefix)
                        {
                            if (!string.IsNullOrEmpty(existingMark))
                            {
                                // Extract the prefix from the existing mark
                                string extractedProjectPrefix = ExtractProjectPrefixFromMark(existingMark, candidatePrefixes);
                                if (!string.IsNullOrEmpty(extractedProjectPrefix))
                                {
                                    effectiveProjectPrefix = extractedProjectPrefix;
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        File.AppendAllText(mepmarkLogPath,
                                            $"[MARK-ASSIGN] Preserving existing project prefix '{effectiveProjectPrefix}' from mark '{existingMark}' (RemarkProjectPrefix=false)\n");
                                    }
                                }
                                else
                                {
                                    // Fall back to the stored project prefix from UI/service
                                    var storedPrefix = markPrefixes.ProjectPrefix;
                                    if (!string.IsNullOrEmpty(storedPrefix))
                                    {
                                        effectiveProjectPrefix = storedPrefix;
                                    }
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        File.AppendAllText(mepmarkLogPath,
                                            $"[MARK-ASSIGN] Existing mark had no project prefix; using stored prefix '{effectiveProjectPrefix}' (RemarkProjectPrefix=false)\n");
                                    }
                                }
                            }
                            else
                            {
                                // No existing mark — use the stored project prefix if available
                                var storedPrefix = markPrefixes.ProjectPrefix;
                                if (!string.IsNullOrEmpty(storedPrefix))
                                {
                                    effectiveProjectPrefix = storedPrefix;
                                }
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    File.AppendAllText(mepmarkLogPath,
                                        $"[MARK-ASSIGN] No existing mark found; using stored project prefix '{effectiveProjectPrefix}' (RemarkProjectPrefix=false)\n");
                                }
                            }
                        }
                        
                        string fullMarkValue = $"{effectiveProjectPrefix}{markValue}";
                        
                                                // ✅ DEPLOYMENT MODE: Skip file writes
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            File.AppendAllText(mepmarkLogPath, 
                            $"[MARK-ASSIGN] Generating mark: prefix='{effectiveProjectPrefix}', discipline='{elementPrefix}', number={numberToUse} → '{fullMarkValue}'\n");
                        }
                        
                        try
                        {
                        SetMarkParameter(sleeve, fullMarkValue);
                        processedCount++;
                        
                                                if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[MarkParameterService] Applied MEPMARK '{fullMarkValue}' to sleeve {sleeve.Id.IntegerValue}");
                                                // ✅ DEPLOYMENT MODE: Skip file writes
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            File.AppendAllText(mepmarkLogPath, $"[MARK-ASSIGN] ✅ SUCCESS: Applied MEPMARK '{fullMarkValue}' to sleeve {sleeve.Id.IntegerValue}\n");
                        }
                        }
                        catch (Exception setEx)
                        {
                            // Parameter setting failed
                            errorCount++;
                                                        if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Error($"[MarkParameterService] Error setting MEPMARK '{fullMarkValue}' on sleeve {sleeve.Id.IntegerValue}: {setEx.Message}");
                                                        // ✅ DEPLOYMENT MODE: Skip file writes
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                File.AppendAllText(mepmarkLogPath, 
                                $"[MARK-ASSIGN] ❌ FAILED: Could not set '{fullMarkValue}' on sleeve {sleeve.Id.IntegerValue}: {setEx.Message}\n");
                            }
                            // If we used a number but failed to set, remove it from usedNumbers if remark=true
                            if (remarkAll && usedNumbers.Contains(numberToUse))
                            {
                                usedNumbers.Remove(numberToUse);
                            }
                            // Don't adjust actualIndex - it's already incremented or preserved correctly
                        }
                    }
                    catch (Exception ex)
                    {
                        errorCount++;
                        // Don't increment actualIndex here - it's already handled in the try block
                                                if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Error($"[MarkParameterService] Error applying MEPMARK to sleeve {allSleeves[i].Id.IntegerValue}: {ex.Message}");
                                                // ✅ DEPLOYMENT MODE: Skip file writes
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            File.AppendAllText(mepmarkLogPath, 
                            $"[MARK-ASSIGN] ❌ EXCEPTION: Sleeve {allSleeves[i].Id.IntegerValue}: {ex.Message}\n");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[MarkParameterService] Error processing category '{category}': {ex.Message}");
                throw;
            }
            finally
            {
                string mepmarkLogPathFinal = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    File.AppendAllText(mepmarkLogPathFinal, $"MEPMARK session completed: {processedCount} processed, {errorCount} errors\n");
                }
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    File.AppendAllText(mepmarkLogPathFinal, $"===== MEPMARK DEBUG SESSION ENDED {DateTime.Now:yyyy-MM-dd HH:mm:ss} =====\n\n");
                }
            }
            
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
                                        // ✅ DEPLOYMENT MODE: Skip file writes
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        File.AppendAllText(mepmarkLogPath, 
                        $"[INDIVIDUAL-DEBUG] Checking sleeve {sleeve.Id}: MEP_ElementId={mepElementId}\n");
                    }
                    
                    // Find matching clash zone (pass doc context for project-specific paths)
                    var clashZone = GetClashZoneByMepElementId(mepElementId, doc);
                                        // ✅ DEPLOYMENT MODE: Skip file writes
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        File.AppendAllText(mepmarkLogPath, 
                        $"[INDIVIDUAL-DEBUG] Sleeve {sleeve.Id}: ClashZone={clashZone != null}, Category={clashZone?.MepElementCategory}\n");
                    }
                    
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
                                                                // ✅ DEPLOYMENT MODE: Skip file writes
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    File.AppendAllText(mepmarkLogPath, 
                                    $"[INDIVIDUAL-DEBUG] ✅ ACCEPTED: Individual sleeve {sleeve.Id} for category '{category}'\n");
                                }
                            }
                            else
                            {
                                                                // ✅ DEPLOYMENT MODE: Skip file writes
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    File.AppendAllText(mepmarkLogPath, 
                                    $"[INDIVIDUAL-DEBUG] ❌ REJECTED: Category mismatch - XML='{clashZone.MepElementCategory}' vs Requested='{category}'\n");
                                }
                            }
                        }
                        else
                        {
                                                        // ✅ DEPLOYMENT MODE: Skip file writes
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                File.AppendAllText(mepmarkLogPath, 
                                $"[INDIVIDUAL-DEBUG] ❌ REJECTED: This is a cluster sleeve, not individual\n");
                            }
                        }
                    }
                }
            }

            string mepmarkLogPathFinal = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                        // ✅ DEPLOYMENT MODE: Skip file writes
            if (!DeploymentConfiguration.DeploymentMode)
            {
                File.AppendAllText(mepmarkLogPathFinal, 
                $"Found {individualSleeves.Count} individual sleeves for category '{category}'\n");
            }
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

            // ✅ CRITICAL FIX: Find cluster sleeves by ClusterInstanceId instead of MEP_ElementId
            // Cluster sleeves don't have a single MEP_ElementId, so we query clash zones by ClusterInstanceId
            var categorySleeves = new List<FamilyInstance>();

            foreach (var sleeve in allClusterSleeves)
            {
                int sleeveId = sleeve.Id.IntegerValue;
                
                // ✅ NEW APPROACH: Query clash zones by ClusterInstanceId (cluster sleeve's own ID)
                var clashZone = GetClashZoneByClusterInstanceId(sleeveId, category, doc);
                
                string mepmarkLogPathDebug = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                
                if (clashZone != null && clashZone.IsClusterResolved && clashZone.ClusterSleeveInstanceId == sleeveId)
                {
                    // Verify category matches
                    bool categoryMatch = clashZone.MepElementCategory == category;
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        File.AppendAllText(mepmarkLogPathDebug, 
                            $"[CLUSTER-DEBUG] Category check: XML='{clashZone.MepElementCategory}' vs Requested='{category}' → Match={categoryMatch}\n");
                    }
                    
                    if (categoryMatch)
                    {
                        categorySleeves.Add(sleeve);
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            File.AppendAllText(mepmarkLogPathDebug, 
                                $"[CLUSTER-DEBUG] ✅ ACCEPTED: Cluster sleeve {sleeve.Id} for category '{category}' " +
                                $"(ClusterSleeveId={clashZone.ClusterSleeveInstanceId})\n");
                        }
                    }
                    else
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            File.AppendAllText(mepmarkLogPathDebug, 
                                $"[CLUSTER-DEBUG] ❌ REJECTED: Category mismatch - XML='{clashZone.MepElementCategory}' vs Requested='{category}'\n");
                        }
                    }
                }
                else
                {
                    // Not a cluster sleeve for this category
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        File.AppendAllText(mepmarkLogPathDebug,
                            $"REJECT {sleeve.Id}: no cluster clash-zone found for ClusterInstanceId {sleeveId} in category '{category}'\n");
                    }
                }
            }

            string mepmarkLogPathFinal = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                        // ✅ DEPLOYMENT MODE: Skip file writes
            if (!DeploymentConfiguration.DeploymentMode)
            {
                File.AppendAllText(mepmarkLogPathFinal, $"Found {categorySleeves.Count} cluster sleeves for category '{category}'\n");
            }
            return categorySleeves;
        }

        /// <summary>
        /// ✅ NEW: Get clash zones by ClusterInstanceId (for cluster sleeves)
        /// Cluster sleeves don't have a single MEP_ElementId, so we query by ClusterInstanceId
        /// </summary>
        private ClashZone GetClashZoneByClusterInstanceId(int clusterInstanceId, string category, Document? doc = null)
        {
            if (doc == null || clusterInstanceId <= 0)
                return null;

            try
            {
                using (var context = new SleeveDbContext(doc))
                {
                    var repository = new ClashZoneRepository(context);
                    
                    // Query database for clash zones by ClusterInstanceId
                    var zones = repository.GetClashZonesByCategory(category);
                    var foundZone = zones?.FirstOrDefault(z => z.ClusterSleeveInstanceId == clusterInstanceId && z.IsClusterResolved);
                    
                    if (foundZone != null)
                    {
                        string mepmarkLogPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            File.AppendAllText(mepmarkLogPath, 
                                $"[CACHE-LOOKUP] ✓ Found cluster clash zone by ClusterInstanceId={clusterInstanceId}, Category={foundZone.MepElementCategory}\n");
                        }
                        return foundZone;
                    }
                }
            }
            catch (Exception ex)
            {
                string mepmarkLogPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    File.AppendAllText(mepmarkLogPath, 
                        $"[CACHE-LOOKUP] ❌ DB query failed for ClusterInstanceId={clusterInstanceId}: {ex.Message}\n");
                }
            }

            return null;
        }

        /// <summary>
        /// Get clash zone by MEP element ID from XML files
        /// ⚠️ PERFORMANCE: Uses cache to avoid O(n·m) XML deserialization
        /// </summary>
        private ClashZone GetClashZoneByMepElementId(long mepElementId, Document? doc = null)
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
                                // ✅ DEPLOYMENT MODE: Skip file writes
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    File.AppendAllText(mepmarkLogPath, 
                    $"[CACHE-LOOKUP] ✓ Found clash zone for MEP_ID={mepElementId}: Category={clashZone.MepElementCategory}, " +
                    $"IsClusterResolved={clashZone.IsClusterResolved}, ClusterSleeveId={clashZone.ClusterSleeveInstanceId}\n");
                }
                return clashZone;
            }

            // ✅ FALLBACK: If not in cache, query database directly
            if (doc != null && mepElementId > 0)
            {
                try
                {
                    using (var context = new SleeveDbContext(doc))
                    {
                        var repository = new ClashZoneRepository(context);
                        
                        // Query database for clash zone by MEP element ID
                        var allCategories = new[] { "Ducts", "Pipes", "Cable Trays", "Duct Accessories" };
                        foreach (var category in allCategories)
                        {
                            var zones = repository.GetClashZonesByCategory(category);
                            var foundZone = zones?.FirstOrDefault(z => z.MepElementIdValue == mepElementId);
                            
                            if (foundZone != null)
                            {
                                // Add to cache for future lookups
                                if (_clashZoneCache == null)
                                    _clashZoneCache = new Dictionary<long, ClashZone>();
                                
                                _clashZoneCache[mepElementId] = foundZone;
                                
                                string mepmarkLogPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    File.AppendAllText(mepmarkLogPath, 
                                    $"[CACHE-LOOKUP] ✓ Found in DB (not cached): MEP_ID={mepElementId}, Category={foundZone.MepElementCategory}\n");
                                }
                                return foundZone;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    string mepmarkLogPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        File.AppendAllText(mepmarkLogPath, 
                        $"[CACHE-LOOKUP] ❌ DB query failed for MEP_ID={mepElementId}: {ex.Message}\n");
                    }
                }
            }

            string mepmarkLogPathNotFound = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
            File.AppendAllText(mepmarkLogPathNotFound, $"[CACHE-LOOKUP] ❌ NOT FOUND: MEP_ID={mepElementId} (cache size: {_clashZoneCache?.Count ?? 0})\n");
            return null;
        }

        /// <summary>
        /// Initialize clash zone cache from database first, then fall back to XML files (called once)
        /// </summary>
        private void InitializeClashZoneCache(Document? doc = null)
        {
            _clashZoneCache = new Dictionary<long, ClashZone>();

            try
            {
                string mepmarkLogPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                                // ✅ DEPLOYMENT MODE: Skip file writes
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    File.AppendAllText(mepmarkLogPath, $"\n[CACHE-INIT] ===== CACHE INITIALIZATION STARTED =====\n");
                }
                File.AppendAllText(mepmarkLogPath, $"[CACHE-INIT] Document: {(doc != null ? doc.Title : "NULL (using Default path)")}\n");
                
                int totalClashZones = 0;
                
                // ✅ STEP 1: Try loading from database first
                if (doc != null)
                {
                    try
                    {
                        using (var context = new SleeveDbContext(doc))
                        {
                            var repository = new ClashZoneRepository(context);
                            
                            // Load clash zones for all MEP categories
                            var allCategories = new[] { "Ducts", "Pipes", "Cable Trays", "Duct Accessories" };
                            
                            foreach (var category in allCategories)
                            {
                                var dbZones = repository.GetClashZonesByCategory(category);
                                
                                if (dbZones != null && dbZones.Count > 0)
                                {
                                    int categoryCount = 0;
                                    foreach (var clashZone in dbZones)
                                    {
                                        if (clashZone != null && clashZone.MepElementIdValue > 0)
                                        {
                                            long key = clashZone.MepElementIdValue;
                                            if (!_clashZoneCache.ContainsKey(key))
                                            {
                                                clashZone.EnsureSleevePlacementPointReconstructed();
                                                _clashZoneCache[key] = clashZone;
                                                totalClashZones++;
                                                categoryCount++;
                                            }
                                        }
                                    }
                                    
                                    if (categoryCount > 0)
                                    {
                                        File.AppendAllText(mepmarkLogPath, 
                                            $"[CACHE-INIT] ✓ Loaded {categoryCount} clash zones from database for category '{category}'\n");
                                    }
                                }
                            }
                            
                            if (totalClashZones > 0)
                            {
                                File.AppendAllText(mepmarkLogPath, 
                                    $"[CACHE-INIT] ✓ Total {totalClashZones} clash zones loaded from database\n");
                                
                                // ✅ DEPLOYMENT MODE: Skip file writes
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    File.AppendAllText(mepmarkLogPath, 
                                        $"[CACHE-INIT] Cache size: {_clashZoneCache.Count} entries\n\n");
                                }
                                return; // Successfully loaded from database, skip XML fallback
                            }
                        }
                    }
                    catch (Exception dbEx)
                    {
                        File.AppendAllText(mepmarkLogPath, 
                            $"[CACHE-INIT] ⚠️ Database load failed: {dbEx.Message}, falling back to XML\n");
                    }
                }
                
                // ✅ STEP 2: Fallback to XML files if database is empty or failed
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
                
                                // ✅ DEPLOYMENT MODE: Skip file writes
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    File.AppendAllText(mepmarkLogPath, $"[CACHE-INIT] Filters Directory: {filtersDirectory}\n");
                }
                
                if (!Directory.Exists(filtersDirectory))
                {
                                        // ✅ DEPLOYMENT MODE: Skip file writes
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        File.AppendAllText(mepmarkLogPath, $"[CACHE-INIT] ❌ ERROR: Filters directory not found: {filtersDirectory}\n");
                    }
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
                
                File.AppendAllText(mepmarkLogPath, $"[CACHE-INIT] Found {allXmlFiles.Length} XML files total, {xmlFiles.Count} filter files (excluded {allXmlFiles.Length - xmlFiles.Count} global/CONDITIONS files)\n");
                
                // ✅ ENHANCED: List all XML files found (including excluded ones for debugging)
                                // ✅ DEPLOYMENT MODE: Skip file writes
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    File.AppendAllText(mepmarkLogPath, $"[CACHE-INIT] === ALL XML FILES IN DIRECTORY ===\n");
                }
                foreach (var xmlFile in allXmlFiles)
                {
                    string fileName = Path.GetFileName(xmlFile);
                    bool isExcluded = fileName.EndsWith("_global.xml", StringComparison.OrdinalIgnoreCase) ||
                                     fileName.EndsWith("_CONDITIONS.xml", StringComparison.OrdinalIgnoreCase) ||
                                     fileName.EndsWith("_conditions.xml", StringComparison.OrdinalIgnoreCase);
                    File.AppendAllText(mepmarkLogPath, $"[CACHE-INIT]   {(isExcluded ? "❌ EXCLUDED" : "✓ INCLUDED")}: {fileName}\n");
                }
                
                                // ✅ DEPLOYMENT MODE: Skip file writes
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    File.AppendAllText(mepmarkLogPath, $"[CACHE-INIT] === PROCESSING {xmlFiles.Count} FILTER FILES ===\n");
                }

                // ✅ Reset counter for XML fallback (database counter already used above)
                totalClashZones = 0;
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
                            
                            // ✅ CRITICAL FIX: Load from BOTH hierarchical structure (Filters → FileCombos → ClashZones) AND flat structure
                            // UpdateSleeveCoordinatesInXml updates both structures, so we must read from both
                            // This ensures floor sleeves stored in hierarchical structure are found!
                            
                            var allClashZonesFromFile = new List<ClashZone>();
                            
                            // 1. Load from hierarchical structure (PRIMARY) - where floor sleeves might be stored
                            if (filter?.ClashZoneStorage?.Filters != null)
                            {
                                int hierarchicalCount = 0;
                                foreach (var filterGroup in filter.ClashZoneStorage.Filters)
                                {
                                    if (filterGroup?.FileCombos != null)
                                    {
                                        foreach (var fileCombo in filterGroup.FileCombos)
                                        {
                                            if (fileCombo?.ClashZones != null)
                                            {
                                                foreach (var cz in fileCombo.ClashZones)
                                                {
                                                    if (cz != null && !allClashZonesFromFile.Any(c => c.Id == cz.Id))
                                                    {
                                                        cz.EnsureSleevePlacementPointReconstructed();
                                                        allClashZonesFromFile.Add(cz);
                                                        hierarchicalCount++;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                
                                if (hierarchicalCount > 0)
                                {
                                    File.AppendAllText(mepmarkLogPath, $"[CACHE-INIT] ✓ Loaded {hierarchicalCount} clash zones from hierarchical structure in {Path.GetFileName(xmlFile)}\n");
                                }
                            }
                            
                            // 2. Load from flat structure (BACKWARD COMPATIBILITY)
                            if (filter?.ClashZoneStorage?.AllZones != null)
                            {
                                int flatCount = 0;
                                foreach (var cz in filter.ClashZoneStorage.AllZones)
                                {
                                    if (cz != null && !allClashZonesFromFile.Any(c => c.Id == cz.Id))
                                    {
                                        cz.EnsureSleevePlacementPointReconstructed();
                                        allClashZonesFromFile.Add(cz);
                                        flatCount++;
                                    }
                                }
                                
                                if (flatCount > 0)
                                {
                                    File.AppendAllText(mepmarkLogPath, $"[CACHE-INIT] ✓ Loaded {flatCount} clash zones from flat structure in {Path.GetFileName(xmlFile)}\n");
                                }
                            }
                            
                            int clashZoneCount = allClashZonesFromFile.Count;
                            if (clashZoneCount > 0)
                            {
                                File.AppendAllText(mepmarkLogPath, $"[CACHE-INIT] ✓ Loading {Path.GetFileName(xmlFile)}: {clashZoneCount} total clash zones found (hierarchical + flat)\n");
                                
                                // ✅ DEBUG: Log first few clash zones to verify they have MEP element IDs
                                if (clashZoneCount > 0)
                                {
                                    var firstZone = allClashZonesFromFile[0];
                                                                        // ✅ DEPLOYMENT MODE: Skip file writes
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        File.AppendAllText(mepmarkLogPath, 
                                        $"[CACHE-INIT]   Sample Zone 0: MEP_ID={firstZone.MepElementId.IntegerValue}, " +
                                        $"Category={firstZone.MepElementCategory}, " +
                                        $"SleeveId={firstZone.SleeveInstanceId}, " +
                                        $"ClusterId={firstZone.ClusterSleeveInstanceId}, " +
                                        $"HostType={firstZone.StructuralElementType}\n");
                                    }
                                }
                                
                                foreach (var clashZone in allClashZonesFromFile)
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
                                                                                        // ✅ DEPLOYMENT MODE: Skip file writes
                                            if (!DeploymentConfiguration.DeploymentMode)
                                            {
                                                File.AppendAllText(mepmarkLogPath, 
                                                $"[CACHE-INIT]   Zone {zonesInFile}: MEP_ID={key}, Category={clashZone.MepElementCategory}, " +
                                                $"IsClusterResolved={clashZone.IsClusterResolved}, ClusterSleeveId={clashZone.ClusterSleeveInstanceId}, " +
                                                $"SleeveId={clashZone.SleeveInstanceId}, HostType={clashZone.StructuralElementType}\n");
                                            }
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
                                                // ✅ DEPLOYMENT MODE: Skip file writes
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            File.AppendAllText(mepmarkLogPath, $"[CACHE-INIT]   Stack: {ex.StackTrace}\n");
                        }
                        continue;
                    }
                }

                                // ✅ DEPLOYMENT MODE: Skip file writes
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    File.AppendAllText(mepmarkLogPath, $"[CACHE-INIT] ===== CACHE INITIALIZATION COMPLETE =====\n");
                }
                                // ✅ DEPLOYMENT MODE: Skip file writes
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    File.AppendAllText(mepmarkLogPath, $"[CACHE-INIT] ✓ Cached {totalClashZones} unique clash zones from {filesProcessed}/{xmlFiles.Count} files\n");
                }
                                // ✅ DEPLOYMENT MODE: Skip file writes
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    File.AppendAllText(mepmarkLogPath, $"[CACHE-INIT] Cache size: {_clashZoneCache.Count} entries\n\n");
                }
            }
            catch (Exception ex)
            {
                string mepmarkLogPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                                // ✅ DEPLOYMENT MODE: Skip file writes
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    File.AppendAllText(mepmarkLogPath, $"[CACHE-INIT] ❌ FATAL ERROR initializing cache: {ex.Message}\n");
                }
                                // ✅ DEPLOYMENT MODE: Skip file writes
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    File.AppendAllText(mepmarkLogPath, $"[CACHE-INIT] Stack: {ex.StackTrace}\n");
                }
                                if (!DeploymentConfiguration.DeploymentMode)
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
                            if (filter?.ClashZoneStorage?.AllZones != null)
                            {
                                foreach (var clashZone in filter.ClashZoneStorage.AllZones)
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

                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[MarkParameterService] No category found for MEP Element ID {mepElementId} in {xmlFiles.Length} XML files");
                    string mepmarkLogPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                                        // ✅ DEPLOYMENT MODE: Skip file writes
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        File.AppendAllText(mepmarkLogPath, $"WARNING: No category found for MEP Element ID {mepElementId} in {xmlFiles.Length} XML files\n");
                    }
                    return "Unknown";
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[MarkParameterService] Error getting category for MEP Element ID {mepElementId}: {ex.Message}");
                return "Unknown";
            }
        }
        
        /// <summary>
        /// ✅ NEW: Get maximum existing mark number per category (any prefix)
        /// Finds max number for sleeves in this category regardless of prefix
        /// </summary>
        private HashSet<string> GetCandidateDisciplinePrefixes(string category, string defaultPrefix, MarkPrefixSettings markPrefixes)
        {
            var prefixes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            if (!string.IsNullOrWhiteSpace(defaultPrefix))
                prefixes.Add(defaultPrefix.Trim());

            if (markPrefixes != null)
            {
                var basePrefix = markPrefixes.GetDisciplinePrefix(category);
                if (!string.IsNullOrWhiteSpace(basePrefix))
                    prefixes.Add(basePrefix.Trim());

                if (category.Equals("Ducts", StringComparison.OrdinalIgnoreCase))
                {
                    foreach (var kvp in markPrefixes.DuctSystemTypeOverrides)
                    {
                        if (!string.IsNullOrWhiteSpace(kvp.Value))
                            prefixes.Add(kvp.Value.Trim());
                    }
                }

                if (category.Equals("Cable Trays", StringComparison.OrdinalIgnoreCase))
                {
                    foreach (var kvp in markPrefixes.CableTrayServiceTypeOverrides)
                    {
                        if (!string.IsNullOrWhiteSpace(kvp.Value))
                            prefixes.Add(kvp.Value.Trim());
                    }
                }
            }

            return prefixes;
        }

        private string ResolveDisciplinePrefixForElement(string category, string defaultPrefix, MarkPrefixSettings markPrefixes, ClashZone clashZone)
        {
            string fallbackPrefix = !string.IsNullOrWhiteSpace(defaultPrefix)
                ? defaultPrefix
                : markPrefixes?.GetDisciplinePrefix(category) ?? defaultPrefix ?? string.Empty;

            if (markPrefixes == null)
                return fallbackPrefix;

            // ✅ CRITICAL: Extract System Type and Service Type from clash zone
            string systemType = GetClashParameterValue(clashZone, "System Type", "MEP System Type", "System Classification");
            string serviceType = GetClashParameterValue(clashZone, "Service Type", "System Abbreviation", "MEP System Type");

            // ✅ DEBUG: Log extracted system/service types and available overrides
            if (!DeploymentConfiguration.DeploymentMode)
            {
                string mepmarkLogPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                File.AppendAllText(mepmarkLogPath, 
                    $"[PREFIX-RESOLVE] Category={category}, SystemType='{systemType ?? "null"}', ServiceType='{serviceType ?? "null"}'\n");
                
                if (category == "Ducts" && markPrefixes.DuctSystemTypeOverrides.Count > 0)
                {
                    var overrideKeys = string.Join(", ", markPrefixes.DuctSystemTypeOverrides.Keys);
                    File.AppendAllText(mepmarkLogPath, 
                        $"[PREFIX-RESOLVE] Available Duct System Type Overrides: {overrideKeys}\n");
                }
                
                if (category == "Cable Trays" && markPrefixes.CableTrayServiceTypeOverrides.Count > 0)
                {
                    var overrideKeys = string.Join(", ", markPrefixes.CableTrayServiceTypeOverrides.Keys);
                    File.AppendAllText(mepmarkLogPath, 
                        $"[PREFIX-RESOLVE] Available Cable Tray Service Type Overrides: {overrideKeys}\n");
                }
            }

            // ✅ CRITICAL: Get prefix using two-tier resolution (System Type override takes precedence)
            var resolved = markPrefixes.GetPrefixForElement(category, systemType, serviceType);
            
            // ✅ DEBUG: Log the resolved prefix
            if (!DeploymentConfiguration.DeploymentMode)
            {
                string mepmarkLogPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                string disciplinePrefix = markPrefixes.GetDisciplinePrefix(category);
                bool isOverride = resolved != disciplinePrefix;
                File.AppendAllText(mepmarkLogPath, 
                    $"[PREFIX-RESOLVE] ✅ Resolved prefix='{resolved}', Discipline prefix='{disciplinePrefix}', IsOverride={isOverride}\n");
            }
            
            if (!string.IsNullOrWhiteSpace(resolved))
                return resolved.Trim();

            return fallbackPrefix;
        }

        private static string GetClashParameterValue(ClashZone clashZone, params string[] keys)
        {
            if (clashZone?.MepParameterValues == null || clashZone.MepParameterValues.Count == 0 || keys == null)
                return null;

            foreach (var key in keys)
            {
                if (string.IsNullOrWhiteSpace(key))
                    continue;

                var param = clashZone.MepParameterValues
                    .FirstOrDefault(kv => kv != null &&
                        kv.Key != null &&
                        kv.Key.Equals(key, StringComparison.OrdinalIgnoreCase));

                if (param != null && !string.IsNullOrWhiteSpace(param.Value))
                    return param.Value;
            }

            return null;
        }

        /// <summary>
        /// ✅ OPTIMIZED: Get max existing mark number for category using database query (10-100× faster than scanning all sleeves)
        /// Uses SleeveSnapshots table joined with ClashZones to get MEP Mark values without scanning Revit elements
        /// Falls back to Revit scan if database query fails or returns no results
        /// </summary>
        private int GetMaxExistingMarkNumberForCategory(Document doc, string category, IEnumerable<string> disciplinePrefixes)
        {
            try
            {
                int maxNumber = 0;
                var prefixSet = new HashSet<string>((disciplinePrefixes ?? Enumerable.Empty<string>()), StringComparer.OrdinalIgnoreCase);
                if (prefixSet.Count == 0)
                    return 0;
                
                // ✅ OPTIMIZATION 1: Check cache first (avoid repeated database queries for same category/prefix)
                string cacheKey = $"{category}_{string.Join("_", prefixSet.OrderBy(p => p))}";
                if (_maxMarkNumberCache != null && _maxMarkCacheDocument == doc && _maxMarkNumberCache.TryGetValue(cacheKey, out int cachedMax))
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[MarkParameterService] ✅ Cache hit: Max existing mark number for category '{category}': {cachedMax}");
                    return cachedMax;
                }
                
                // Initialize cache if needed
                if (_maxMarkNumberCache == null || _maxMarkCacheDocument != doc)
                {
                    _maxMarkNumberCache = new Dictionary<string, int>();
                    _maxMarkCacheDocument = doc;
                }
                
                // ✅ OPTIMIZATION 2: Try database query (much faster than scanning all sleeves)
                int dbMaxNumber = GetMaxMarkNumberFromDatabase(doc, category, prefixSet);
                if (dbMaxNumber > 0)
                {
                    // Cache the result
                    _maxMarkNumberCache[cacheKey] = dbMaxNumber;
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[MarkParameterService] ✅ Database query: Max existing mark number for category '{category}': {dbMaxNumber}");
                    return dbMaxNumber;
                }
                
                // ✅ FALLBACK: If database query returns 0 or fails, scan Revit elements (slower but reliable)
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[MarkParameterService] ⚠️ Database query returned 0, falling back to Revit scan for category '{category}'");
                
                // ✅ OPTIMIZED: Only get sleeve family instances
                var sleeveElements = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(fi => {
                        var famName = fi.Symbol?.Family?.Name ?? string.Empty;
                        return famName.IndexOf("OpeningOnWall", StringComparison.OrdinalIgnoreCase) >= 0 ||
                               famName.IndexOf("OpeningOnSlab", StringComparison.OrdinalIgnoreCase) >= 0;
                    })
                    .ToList();
                
                foreach (var element in sleeveElements)
                {
                    // Get category for this sleeve using MEP_ElementId
                    var mepElementIdParam = element.LookupParameter("MEP_ElementId");
                    if (mepElementIdParam != null)
                    {
                        long mepElementId = mepElementIdParam.AsInteger();
                        var clashZone = GetClashZoneByMepElementId(mepElementId, doc);
                        
                        // Only consider sleeves in the same category
                        if (clashZone?.MepElementCategory != category)
                            continue;
                    }
                    
                    // Resolve mark parameter
                    var markParam = ResolveMarkParameter(element);
                    if (markParam != null)
                    {
                        string markValue = markParam.AsString() ?? "";

                        if (!string.IsNullOrEmpty(markValue))
                        {
                            // Try to extract number from any mark format (any prefix + discipline prefix + number)
                            int? extractedNumber = ExtractNumberFromMark(markValue, prefixSet);
                            if (extractedNumber.HasValue)
                            {
                                maxNumber = Math.Max(maxNumber, extractedNumber.Value);
                            }
                        }
                    }
                }
                
                // Cache the result (even if 0, to avoid repeated scans)
                _maxMarkNumberCache[cacheKey] = maxNumber;
                
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[MarkParameterService] Revit scan: Max existing mark number for category '{category}': {maxNumber}");
                return maxNumber;
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[MarkParameterService] Error getting max existing mark number for category: {ex.Message}");
                return 0; // Start from 1 if error
            }
        }
        
        /// <summary>
        /// ✅ OPTIMIZATION: Query database to get max MEP Mark number for a category (10-100× faster than scanning all sleeves)
        /// Queries SleeveSnapshots table joined with ClashZones to get MEP Mark values from database
        /// Returns 0 if database query fails or no marks found (fallback to Revit scan)
        /// </summary>
        private int GetMaxMarkNumberFromDatabase(Document doc, string category, HashSet<string> disciplinePrefixes)
        {
            try
            {
                using (var dbContext = new SleeveDbContext(doc))
                {
                    using (var cmd = dbContext.Connection.CreateCommand())
                    {
                        // ✅ OPTIMIZED QUERY: Join SleeveSnapshots with ClashZones to filter by category and get MEP Mark values
                        // This avoids scanning all Revit elements - database query is 10-100× faster
                        cmd.CommandText = @"
                            SELECT DISTINCT ss.MepParametersJson, cz.MepElementCategory
                            FROM SleeveSnapshots ss
                            INNER JOIN ClashZones cz ON (
                                (ss.SleeveInstanceId IS NOT NULL AND ss.SleeveInstanceId = cz.SleeveInstanceId) OR
                                (ss.ClusterInstanceId IS NOT NULL AND ss.ClusterInstanceId = cz.ClusterInstanceId)
                            )
                            WHERE cz.MepElementCategory = @Category
                              AND ss.MepParametersJson IS NOT NULL
                              AND ss.MepParametersJson != '{}'
                              AND ss.MepParametersJson != ''";
                        
                        cmd.Parameters.AddWithValue("@Category", category);
                        
                        int maxNumber = 0;
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                try
                                {
                                    var mepParamsJson = reader.IsDBNull(0) ? null : reader.GetString(0);
                                    if (string.IsNullOrWhiteSpace(mepParamsJson))
                                        continue;
                                    
                                    // Parse JSON to get MEP Mark value
                                    var mepParams = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, string>>(mepParamsJson);
                                    if (mepParams != null && mepParams.TryGetValue("MEP Mark", out var markValue) && !string.IsNullOrWhiteSpace(markValue))
                                    {
                                        // Try alternative parameter names
                                        if (string.IsNullOrWhiteSpace(markValue) && mepParams.TryGetValue("Mark", out markValue))
                                        {
                                            // Use Mark if MEP Mark not found
                                        }
                                        
                                        if (!string.IsNullOrWhiteSpace(markValue))
                                        {
                                            // Extract number from mark value
                                            int? extractedNumber = ExtractNumberFromMark(markValue, disciplinePrefixes);
                                            if (extractedNumber.HasValue)
                                            {
                                                maxNumber = Math.Max(maxNumber, extractedNumber.Value);
                                            }
                                        }
                                    }
                                }
                                catch (Exception parseEx)
                                {
                                    // Skip invalid JSON entries
                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Warning($"[MarkParameterService] Error parsing MEP parameters JSON: {parseEx.Message}");
                                    continue;
                                }
                            }
                        }
                        
                        return maxNumber;
                    }
                }
            }
            catch (Exception ex)
            {
                // Database query failed - fall back to Revit scan
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[MarkParameterService] Database query for max mark number failed: {ex.Message}, falling back to Revit scan");
                return 0;
            }
        }
        
        /// <summary>
        /// Extract project prefix from existing mark value
        /// Returns everything before the discipline prefix
        /// Example: "PROJ_DCT001" → "PROJ_", "OLD_PRE_PLU002" → "OLD_PRE_", "PLU003" → ""
        /// </summary>
        private string ExtractProjectPrefixFromMark(string markValue, IEnumerable<string> disciplinePrefixes)
        {
            try
            {
                if (string.IsNullOrEmpty(markValue))
                    return string.Empty;

                var prefixes = (disciplinePrefixes ?? Enumerable.Empty<string>())
                    .Where(p => !string.IsNullOrWhiteSpace(p))
                    .OrderByDescending(p => p.Length)
                    .ToList();

                foreach (var prefix in prefixes)
                {
                    int index = markValue.LastIndexOf(prefix, StringComparison.OrdinalIgnoreCase);
                    if (index >= 0)
                    {
                        return markValue.Substring(0, index);
                    }
                }

                return string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }
        
        /// <summary>
        /// ✅ NEW: Extract number from existing mark value
        /// Supports formats like: "PREFIX_DCT001", "OLD_PRE_DCT002", "DCT003", etc.
        /// </summary>
        private int? ExtractNumberFromMark(string markValue, IEnumerable<string> disciplinePrefixes)
        {
            try
            {
                if (string.IsNullOrEmpty(markValue))
                    return null;

                var prefixes = (disciplinePrefixes ?? Enumerable.Empty<string>())
                    .Where(p => !string.IsNullOrWhiteSpace(p))
                    .OrderByDescending(p => p.Length)
                    .ToList();

                foreach (var prefix in prefixes)
                {
                    int prefixIndex = markValue.LastIndexOf(prefix, StringComparison.OrdinalIgnoreCase);
                    if (prefixIndex < 0)
                        continue;

                    string numberPart = markValue.Substring(prefixIndex + prefix.Length);
                    if (int.TryParse(numberPart, out int number))
                    {
                        return number;
                    }
                }

                // Alternative: Look for any trailing digits (last 1-4 digits)
                string digitsOnly = new string(markValue.Reverse().TakeWhile(char.IsDigit).Reverse().ToArray());
                if (!string.IsNullOrEmpty(digitsOnly) && int.TryParse(digitsOnly, out int trailingNumber))
                {
                    return trailingNumber;
                }

                return null;
            }
            catch
            {
                return null;
            }
        }
        
        /// <summary>
        /// Generate mark value based on discipline prefix and number with specified format
        /// </summary>
        private string GenerateMarkValue(string disciplinePrefix, int number, string numberFormat = "000")
        {
            // Calculate padding length from format string (e.g., "00" = 2, "000" = 3, "0000" = 4)
            int paddingLength = numberFormat.Length;
            return $"{disciplinePrefix}{number.ToString().PadLeft(paddingLength, '0')}";
        }
        
        /// <summary>
        /// Set MEP Mark parameter on element
        /// </summary>
        private void SetMarkParameter(Element element, string markValue)
        {
            // Resolve parameter name case/space/underscore-insensitively
            var markParam = ResolveMarkParameter(element);
            string mepmarkLogPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");

                        // ✅ DEPLOYMENT MODE: Skip file writes
            if (!DeploymentConfiguration.DeploymentMode)
            {
                File.AppendAllText(mepmarkLogPath, 
                $"[SET-DEBUG] Element {element.Id}: ResolvedParam='{markParam?.Definition?.Name}', Value='{markValue}', Doc.IsModifiable={element.Document.IsModifiable}\n");
            }
            
            if (markParam == null)
            {
                // ✅ ENHANCED: Log all available parameters for debugging
                var allParams = string.Join(", ", element.Parameters.Cast<Parameter>().Select(p => $"'{p.Definition?.Name}'"));
                                // ✅ DEPLOYMENT MODE: Skip file writes
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    File.AppendAllText(mepmarkLogPath, 
                    $"[SET-FAIL] Element {element.Id}: MEP Mark parameter not found. Available parameters: {allParams}\n");
                }
                throw new InvalidOperationException($"MEP Mark parameter not found on element {element.Id.IntegerValue}. Available params: {allParams}");
            }
            
            if (markParam.IsReadOnly)
            {
                                // ✅ DEPLOYMENT MODE: Skip file writes
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    File.AppendAllText(mepmarkLogPath, 
                    $"[SET-FAIL] Element {element.Id}: Parameter '{markParam.Definition?.Name}' is read-only\n");
                }
                throw new InvalidOperationException($"MEP Mark parameter '{markParam.Definition?.Name}' is read-only on element {element.Id.IntegerValue}");
            }
            
            if (markParam.StorageType != StorageType.String)
            {
                File.AppendAllText(mepmarkLogPath, 
                    $"[SET-FAIL] Element {element.Id}: Parameter '{markParam.Definition?.Name}' has wrong storage type: {markParam.StorageType} (expected String)\n");
                throw new InvalidOperationException($"MEP Mark parameter '{markParam.Definition?.Name}' has wrong storage type {markParam.StorageType} on element {element.Id.IntegerValue}");
            }
            
            try
            {
                markParam.Set(markValue);
                
                // ✅ FIX: Try reading back immediately
                var readBack = markParam.AsString();
                
                // ✅ ENHANCED: More lenient verification - check if value is set (even if not exact match due to whitespace)
                var normalizedReadBack = (readBack ?? string.Empty).Trim();
                var normalizedMarkValue = markValue.Trim();
                
                if (string.Equals(normalizedReadBack, normalizedMarkValue, StringComparison.Ordinal))
                {
                                        // ✅ DEPLOYMENT MODE: Skip file writes
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        File.AppendAllText(mepmarkLogPath, 
                        $"[SET-SUCCESS] Applied '{markValue}' to element {element.Id} using '{markParam.Definition?.Name}'\n");
                    }

                }
                else if (string.IsNullOrEmpty(normalizedReadBack))
                {
                    // ✅ FIX: If read-back is empty, check if parameter needs to be loaded as shared parameter
                    File.AppendAllText(mepmarkLogPath, 
                        $"[SET-VERIFY-FAIL] Element {element.Id}: attempted '{markValue}', read-back is empty (may need shared parameter loaded)\n");
                    
                    // Check if this might be a shared parameter issue
                    if (markParam.Definition != null)
                    {
                        if (markParam.Definition is ExternalDefinition externalDef)
                        {
                            var guid = externalDef.GUID;
                                                        // ✅ DEPLOYMENT MODE: Skip file writes
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                File.AppendAllText(mepmarkLogPath, 
                                $"[SET-VERIFY-FAIL] Parameter GUID: {guid}, IsShared: true\n");
                            }
                        }
                        else if (markParam.Definition is InternalDefinition internalDef)
                        {
                            File.AppendAllText(mepmarkLogPath, 
                                $"[SET-VERIFY-FAIL] Parameter is Internal (built-in), IsShared: {internalDef.VariesAcrossGroups}\n");
                        }
                    }
                    
                    // ✅ FIX: Don't throw exception for empty read-back - parameter might be valid but not readable immediately
                    // Instead, log warning and continue (transaction commit might fix it)
                                        // ✅ DEPLOYMENT MODE: Skip file writes
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        File.AppendAllText(mepmarkLogPath, 
                        $"[SET-WARNING] Parameter set but read-back empty - will verify after transaction commit\n");
                    }
                }
                else
                {
                                        // ✅ DEPLOYMENT MODE: Skip file writes
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        File.AppendAllText(mepmarkLogPath, 
                        $"[SET-VERIFY-FAIL] Element {element.Id}: attempted '{markValue}', read-back='{readBack}'\n");
                    }
                    throw new InvalidOperationException($"MEP Mark write did not persist on element {element.Id.IntegerValue}: expected '{markValue}', got '{readBack}'");
                }
            }
            catch (Exception setEx)
            {
                                // ✅ DEPLOYMENT MODE: Skip file writes
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    File.AppendAllText(mepmarkLogPath, 
                    $"[SET-EXCEPTION] Element {element.Id}: Exception setting parameter: {setEx.Message}\n");
                }
                throw;
            }
        }

        private Parameter ResolveMarkParameter(Element element)
        {
            // Accept: "MEP Mark", "MEP_Mark", "mep mark", "mep_mark" (case-insensitive), or fallback "Mark"
            var preferredNames = new[] { "mepmark", "mep_mark", "mep mark" };
            var fallbackNames = new[] { "mark" };

            try
            {
                // ✅ ENHANCED: Collect all parameters for debugging
                var allParams = element.Parameters.Cast<Parameter>().ToList();
                string mepmarkLogPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                
                // Search all instance parameters case/space/underscore-insensitively
                foreach (Parameter p in allParams)
                {
                    var defName = p.Definition?.Name ?? string.Empty;
                    var norm = NormalizeName(defName);
                    if (preferredNames.Any(n => NormalizeName(n) == norm))
                    {
                        File.AppendAllText(mepmarkLogPath, 
                            $"[PARAM-RESOLVE] Element {element.Id}: Found preferred param '{defName}' (normalized: '{norm}')\n");
                        return p;
                    }
                }
                // Fallback to simple "Mark"
                foreach (Parameter p in allParams)
                {
                    var defName = p.Definition?.Name ?? string.Empty;
                    var norm = NormalizeName(defName);
                    if (fallbackNames.Any(n => NormalizeName(n) == norm))
                    {
                        File.AppendAllText(mepmarkLogPath, 
                            $"[PARAM-RESOLVE] Element {element.Id}: Found fallback param '{defName}' (normalized: '{norm}')\n");
                        return p;
                    }
                }
                
                // ✅ DEBUG: Log all available parameters if not found
                var paramNames = string.Join(", ", allParams.Select(p => $"'{p.Definition?.Name}' ({p.StorageType}, readonly={p.IsReadOnly})"));
                                // ✅ DEPLOYMENT MODE: Skip file writes
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    File.AppendAllText(mepmarkLogPath, 
                    $"[PARAM-RESOLVE] Element {element.Id}: No MEP Mark parameter found. Available params: {paramNames}\n");
                }
            }
            catch (Exception ex)
            {
                string mepmarkLogPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                                // ✅ DEPLOYMENT MODE: Skip file writes
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    File.AppendAllText(mepmarkLogPath, 
                    $"[PARAM-RESOLVE] Element {element.Id}: Exception resolving parameter: {ex.Message}\n");
                }
            }
            
            // Final fallback: direct lookups (in case)
            var param = element.LookupParameter("MEP Mark") ?? element.LookupParameter("MEP_Mark") ?? element.LookupParameter("Mark");
            if (param != null)
            {
                string mepmarkLogPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                                // ✅ DEPLOYMENT MODE: Skip file writes
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    File.AppendAllText(mepmarkLogPath, 
                    $"[PARAM-RESOLVE] Element {element.Id}: Found via direct lookup: '{param.Definition?.Name}'\n");
                }
            }
            return param;
        }

        private string NormalizeName(string s)
        {
            return (s ?? string.Empty).ToLowerInvariant().Replace(" ", string.Empty).Replace("_", string.Empty).Trim();
        }
    }
}

