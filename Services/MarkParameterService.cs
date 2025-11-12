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

        public (int processedCount, int errorCount) ApplyMepMarkToClusters(
            Document doc, string category, string projectPrefix, string disciplinePrefix, bool remarkAll = false, string numberFormat = "000")
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
                
                // ✅ FIX: Get max number per category (not per prefix) - finds max for any prefix in this category
                int categoryMaxNumber = GetMaxExistingMarkNumberForCategory(doc, category, disciplinePrefix);
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
                        
                        // ✅ ENHANCED: Check if MEP Mark exists
                        var existingMark = sleeve.LookupParameter("MEP Mark")?.AsString() ?? 
                                          sleeve.LookupParameter("Mark")?.AsString();
                        
                        int numberToUse;
                        
                        // ✅ FIX: When remark=true, extract number from existing mark (preserve number, update prefix only)
                        if (remarkAll && !string.IsNullOrEmpty(existingMark))
                        {
                            // Try to extract number from existing mark
                            int? extractedNumber = ExtractNumberFromMark(existingMark, disciplinePrefix);
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
                        var mepElementIdParam = sleeve.LookupParameter("MEP_ElementId");
                        long mepElementId = mepElementIdParam?.AsInteger() ?? -1;
                        var clashZone = GetClashZoneByMepElementId(mepElementId, doc);
                        
                        // ✅ DEBUG: Log sleeve host type
                        var famName = sleeve.Symbol?.Family?.Name ?? "";
                        var hostType = famName.IndexOf("OpeningOnWall", StringComparison.OrdinalIgnoreCase) >= 0 ? "Wall" :
                                      famName.IndexOf("OpeningOnSlab", StringComparison.OrdinalIgnoreCase) >= 0 ? "Floor" : "Unknown";
                        
                        if (!DeploymentConfiguration.DeploymentMode && i < 10) // Log first 10 sleeves
                        {
                            File.AppendAllText(mepmarkLogPath, 
                                $"[MARK-PROCESS] Sleeve {sleeve.Id}: HostType={hostType}, Family={famName}, MEP_ID={mepElementId}, " +
                                $"Category={clashZone?.MepElementCategory ?? "null"}, RemarkAll={remarkAll}, ExistingMark='{existingMark ?? "null"}'\n");
                        }
                        
                                                // ✅ DEPLOYMENT MODE: Skip file writes
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            File.AppendAllText(mepmarkLogPath, 
                            $"[MARK-ASSIGN] Sleeve {sleeve.Id}: MEP_ID={mepElementId}, Category={clashZone?.MepElementCategory ?? "UNKNOWN"}, " +
                            $"IsCluster={clashZone?.IsClusterResolved ?? false}, ClusterId={clashZone?.ClusterSleeveInstanceId ?? -1}\n");
                        }
                        
                        string markValue = GenerateMarkValue(disciplinePrefix, numberToUse, numberFormat);
                        string fullMarkValue = $"{projectPrefix}{markValue}";
                        
                                                // ✅ DEPLOYMENT MODE: Skip file writes
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            File.AppendAllText(mepmarkLogPath, 
                            $"[MARK-ASSIGN] Generating mark: prefix='{projectPrefix}', discipline='{disciplinePrefix}', number={numberToUse} → '{fullMarkValue}'\n");
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
                                                // ✅ DEPLOYMENT MODE: Skip file writes
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            File.AppendAllText(mepmarkLogPathDebug,
                            $"REJECT {sleeve.Id}: no clash-zone for MEPid {mepElementId}\n");
                        }
                    else if (!clashZone.IsClusterResolved)
                                                // ✅ DEPLOYMENT MODE: Skip file writes
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            File.AppendAllText(mepmarkLogPathDebug,
                            $"REJECT {sleeve.Id}: IsClusterResolved=false\n");
                        }
                    else if (clashZone.ClusterSleeveInstanceId != sleeve.Id.IntegerValue)
                                                // ✅ DEPLOYMENT MODE: Skip file writes
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            File.AppendAllText(mepmarkLogPathDebug,
                            $"REJECT {sleeve.Id}: ClusterSleeveInstanceId={clashZone.ClusterSleeveInstanceId} != {sleeve.Id.IntegerValue}\n");
                        }
                    else if (clashZone.MepElementCategory != category)
                                                // ✅ DEPLOYMENT MODE: Skip file writes
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            File.AppendAllText(mepmarkLogPathDebug,
                            $"REJECT {sleeve.Id}: category='{clashZone.MepElementCategory}' != '{category}'\n");
                        }
                    
                    if (clashZone != null && clashZone.IsClusterResolved && clashZone.ClusterSleeveInstanceId > 0)
                    {
                        // Check if this sleeve is the actual cluster sleeve
                        if (sleeve.Id.IntegerValue == clashZone.ClusterSleeveInstanceId)
                        {
                            // Verify category matches
                            bool categoryMatch = clashZone.MepElementCategory == category;
                                                        // ✅ DEPLOYMENT MODE: Skip file writes
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                File.AppendAllText(mepmarkLogPathDebug, 
                                $"[CLUSTER-DEBUG] Category check: XML='{clashZone.MepElementCategory}' vs Requested='{category}' → Match={categoryMatch}\n");
                            }
                            
                            if (categoryMatch)
                            {
                                categorySleeves.Add(sleeve);
                                File.AppendAllText(mepmarkLogPathDebug, 
                                    $"[CLUSTER-DEBUG] ✅ ACCEPTED: Cluster sleeve {sleeve.Id} for category '{category}' " +
                                    $"(ClusterSleeveId={clashZone.ClusterSleeveInstanceId}, MEP_ID={mepElementId})\n");
                            }
                            else
                            {
                                                                // ✅ DEPLOYMENT MODE: Skip file writes
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    File.AppendAllText(mepmarkLogPathDebug, 
                                    $"[CLUSTER-DEBUG] ❌ REJECTED: Category mismatch - XML='{clashZone.MepElementCategory}' vs Requested='{category}'\n");
                                }
                            }
                        }
                    }
                }
                else
                {
                    string mepmarkLogPathMissing = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                                        // ✅ DEPLOYMENT MODE: Skip file writes
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        File.AppendAllText(mepmarkLogPathMissing, $"Sleeve {sleeve.Id} missing MEP_ElementId parameter\n");
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
                                // ✅ DEPLOYMENT MODE: Skip file writes
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    File.AppendAllText(mepmarkLogPath, 
                    $"[CACHE-LOOKUP] ✓ Found clash zone for MEP_ID={mepElementId}: Category={clashZone.MepElementCategory}, " +
                    $"IsClusterResolved={clashZone.IsClusterResolved}, ClusterSleeveId={clashZone.ClusterSleeveInstanceId}\n");
                }
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
                                // ✅ DEPLOYMENT MODE: Skip file writes
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    File.AppendAllText(mepmarkLogPath, $"\n[CACHE-INIT] ===== CACHE INITIALIZATION STARTED =====\n");
                }
                File.AppendAllText(mepmarkLogPath, $"[CACHE-INIT] Document: {(doc != null ? doc.Title : "NULL (using Default path)")}\n");
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
                            if (filter?.ClashZoneStorage?.AllZones != null)
                            {
                                int clashZoneCount = filter.ClashZoneStorage.AllZones.Count;
                                File.AppendAllText(mepmarkLogPath, $"[CACHE-INIT] ✓ Loading {Path.GetFileName(xmlFile)}: {clashZoneCount} clash zones found\n");
                                
                                // ✅ DEBUG: Log first few clash zones to verify they have MEP element IDs
                                if (clashZoneCount > 0)
                                {
                                    var firstZone = filter.ClashZoneStorage.AllZones[0];
                                                                        // ✅ DEPLOYMENT MODE: Skip file writes
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        File.AppendAllText(mepmarkLogPath, 
                                        $"[CACHE-INIT]   Sample Zone 0: MEP_ID={firstZone.MepElementId.IntegerValue}, " +
                                        $"Category={firstZone.MepElementCategory}, " +
                                        $"SleeveId={firstZone.SleeveInstanceId}, " +
                                        $"ClusterId={firstZone.ClusterSleeveInstanceId}\n");
                                    }
                                }
                                
                                foreach (var clashZone in filter.ClashZoneStorage.AllZones)
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
                                                $"SleeveId={clashZone.SleeveInstanceId}\n");
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
        private int GetMaxExistingMarkNumberForCategory(Document doc, string category, string disciplinePrefix)
        {
            try
            {
                int maxNumber = 0;
                
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
                            int? extractedNumber = ExtractNumberFromMark(markValue, disciplinePrefix);
                            if (extractedNumber.HasValue)
                            {
                                maxNumber = Math.Max(maxNumber, extractedNumber.Value);
                            }
                        }
                    }
                }
                
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[MarkParameterService] Max existing mark number for category '{category}': {maxNumber}");
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
        /// ✅ NEW: Extract number from existing mark value
        /// Supports formats like: "PREFIX_DCT001", "OLD_PRE_DCT002", "DCT003", etc.
        /// </summary>
        private int? ExtractNumberFromMark(string markValue, string disciplinePrefix)
        {
            try
            {
                if (string.IsNullOrEmpty(markValue) || string.IsNullOrEmpty(disciplinePrefix))
                    return null;
                
                // Look for discipline prefix in the mark
                int prefixIndex = markValue.LastIndexOf(disciplinePrefix, StringComparison.OrdinalIgnoreCase);
                if (prefixIndex < 0)
                    return null;
                
                // Get the part after the discipline prefix
                string numberPart = markValue.Substring(prefixIndex + disciplinePrefix.Length);
                
                // Try to parse as integer (handles formats like "001", "002", "03", etc.)
                if (int.TryParse(numberPart, out int number))
                {
                    return number;
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

