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
                // Path to shared parameter file
                string sharedParamFile = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Resources\Opening family shared parameter.txt";

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
            Document doc, string category, string projectPrefix, string disciplinePrefix, bool remarkAll = false, string numberFormat = "000")
        {
            int processedCount = 0;
            int errorCount = 0;

            try
            {
                // Initialize dedicated MEPMARK debug log
                string mepmarkLogPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\mepmark_debug.log";
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
                        
                        string markValue = GenerateMarkValue(disciplinePrefix, actualIndex, numberFormat);
                        
                        // ✅ Only add project prefix if it's not blank
                        string fullMarkValue = string.IsNullOrWhiteSpace(projectPrefix) 
                            ? markValue 
                            : $"{projectPrefix}{markValue}";
                        
                        SetMarkParameter(sleeve, fullMarkValue);
                        processedCount++;
                        actualIndex++;
                        
                        DebugLogger.Info($"[MarkParameterService] Applied MEPMARK '{fullMarkValue}' to sleeve {sleeve.Id.IntegerValue}");
                        File.AppendAllText(mepmarkLogPath, $"Applied MEPMARK '{fullMarkValue}' to sleeve {sleeve.Id.IntegerValue}\n");
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
            
            File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\mepmark_debug.log", $"MEPMARK session completed: {processedCount} processed, {errorCount} errors\n");
            File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\mepmark_debug.log", $"===== MEPMARK DEBUG SESSION ENDED {DateTime.Now:yyyy-MM-dd HH:mm:ss} =====\n\n");
            
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
                    File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\mepmark_debug.log", 
                        $"[INDIVIDUAL-DEBUG] Checking sleeve {sleeve.Id}: MEP_ElementId={mepElementId}\n");
                    
                    // Find matching clash zone
                    var clashZone = GetClashZoneByMepElementId(mepElementId);
                    File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\mepmark_debug.log", 
                        $"[INDIVIDUAL-DEBUG] Sleeve {sleeve.Id}: ClashZone={clashZone != null}, Category={clashZone?.MepElementCategory}\n");
                    
                    if (clashZone != null && clashZone.MepElementCategory == category)
                    {
                        // ✅ Individual sleeve = NOT part of a cluster (check ClusterSleeveInstanceId)
                        // If this sleeve's ID doesn't match any ClusterSleeveInstanceId, it's an individual sleeve
                        bool isClusterSleeve = (clashZone.IsClusterResolved && 
                                               clashZone.ClusterSleeveInstanceId > 0 && 
                                               sleeve.Id.IntegerValue == clashZone.ClusterSleeveInstanceId);
                        
                        File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\mepmark_debug.log", 
                            $"[INDIVIDUAL-DEBUG] Sleeve {sleeve.Id}: isClusterSleeve={isClusterSleeve} (IsClusterResolved={clashZone.IsClusterResolved}, ClusterSleeveInstanceId={clashZone.ClusterSleeveInstanceId})\n");
                        
                        if (!isClusterSleeve)
                        {
                            individualSleeves.Add(sleeve);
                            File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\mepmark_debug.log", 
                                $"✓ Found individual sleeve {sleeve.Id} for category '{category}'\n");
                        }
                    }
                }
            }

            File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\mepmark_debug.log", 
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

            File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\mepmark_debug.log", 
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

            File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\mepmark_debug.log", $"Found {allClusterSleeves.Count} total opening sleeves (after family name filter)\n");

            // ✅ NEW APPROACH: Find cluster sleeves using IsClusterResolved flag
            var categorySleeves = new List<FamilyInstance>();

            foreach (var sleeve in allClusterSleeves)
            {
                var mepElementIdParam = sleeve.LookupParameter("MEP_ElementId");
                if (mepElementIdParam != null)
                {
                    long mepElementId = mepElementIdParam.AsInteger();
                    
                    // Find matching clash zone in XML files
                    var clashZone = GetClashZoneByMepElementId(mepElementId);
                    
                    // ⚠️ DIAGNOSTIC: Log exactly why sleeves are rejected
                    if (clashZone == null)
                        File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\mepmark_debug.log",
                            $"REJECT {sleeve.Id}: no clash-zone for MEPid {mepElementId}\n");
                    else if (!clashZone.IsClusterResolved)
                        File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\mepmark_debug.log",
                            $"REJECT {sleeve.Id}: IsClusterResolved=false\n");
                    else if (clashZone.ClusterSleeveInstanceId != sleeve.Id.IntegerValue)
                        File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\mepmark_debug.log",
                            $"REJECT {sleeve.Id}: ClusterSleeveInstanceId={clashZone.ClusterSleeveInstanceId} != {sleeve.Id.IntegerValue}\n");
                    else if (clashZone.MepElementCategory != category)
                        File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\mepmark_debug.log",
                            $"REJECT {sleeve.Id}: category='{clashZone.MepElementCategory}' != '{category}'\n");
                    
                    if (clashZone != null && clashZone.IsClusterResolved && clashZone.ClusterSleeveInstanceId > 0)
                    {
                        // Check if this sleeve is the actual cluster sleeve
                        if (sleeve.Id.IntegerValue == clashZone.ClusterSleeveInstanceId)
                        {
                            // Verify category matches
                            if (clashZone.MepElementCategory == category)
                            {
                                categorySleeves.Add(sleeve);
                                File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\mepmark_debug.log", $"✓ Found cluster sleeve {sleeve.Id} for category '{category}' (cluster sleeve ID: {clashZone.ClusterSleeveInstanceId})\n");
                            }
                        }
                    }
                }
                else
                {
                    File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\mepmark_debug.log", $"Sleeve {sleeve.Id} missing MEP_ElementId parameter\n");
                }
            }

            File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\mepmark_debug.log", $"Found {categorySleeves.Count} cluster sleeves for category '{category}'\n");
            return categorySleeves;
        }

        /// <summary>
        /// Get clash zone by MEP element ID from XML files
        /// ⚠️ PERFORMANCE: Uses cache to avoid O(n·m) XML deserialization
        /// </summary>
        private ClashZone GetClashZoneByMepElementId(long mepElementId)
        {
            // Initialize cache once per service instance
            if (!_cacheInitialized)
            {
                InitializeClashZoneCache();
                _cacheInitialized = true;
            }

            // Fast O(1) lookup
            if (_clashZoneCache != null && _clashZoneCache.TryGetValue(mepElementId, out ClashZone clashZone))
            {
                return clashZone;
            }

            return null;
        }

        /// <summary>
        /// Initialize clash zone cache from all XML files (called once)
        /// </summary>
        private void InitializeClashZoneCache()
        {
            _clashZoneCache = new Dictionary<long, ClashZone>();

            try
            {
                var filtersDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings", "Projects", "Default", "Filters");
                
                if (!Directory.Exists(filtersDirectory))
                {
                    File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\mepmark_debug.log", $"[CACHE] Filters directory not found: {filtersDirectory}\n");
                    return;
                }

                var xmlFiles = Directory.GetFiles(filtersDirectory, "*.xml");
                File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\mepmark_debug.log", $"[CACHE] Loading {xmlFiles.Length} XML files into cache...\n");

                int totalClashZones = 0;
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
                                    long key = clashZone.MepElementId.IntegerValue;
                                    if (!_clashZoneCache.ContainsKey(key))
                                    {
                                        _clashZoneCache[key] = clashZone;
                                        totalClashZones++;
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\mepmark_debug.log", $"[CACHE] Error loading {Path.GetFileName(xmlFile)}: {ex.Message}\n");
                        continue;
                    }
                }

                File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\mepmark_debug.log", $"[CACHE] ✓ Cached {totalClashZones} clash zones from {xmlFiles.Length} files\n");
            }
            catch (Exception ex)
            {
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
                    File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\mepmark_debug.log", $"WARNING: No category found for MEP Element ID {mepElementId} in {xmlFiles.Length} XML files\n");
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
                // ✅ Only include project prefix if it's not blank
                string expectedPrefix = string.IsNullOrWhiteSpace(projectPrefix)
                    ? disciplinePrefix
                    : $"{projectPrefix}{disciplinePrefix}";
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
        /// Generate mark value based on discipline prefix, number, and format
        /// </summary>
        private string GenerateMarkValue(string disciplinePrefix, int number, string numberFormat = "000")
        {
            // Format number based on user selection: "00", "000", or "0000"
            string formattedNumber = numberFormat switch
            {
                "00" => $"{number:00}",
                "000" => $"{number:000}",
                "0000" => $"{number:0000}",
                _ => $"{number:000}" // Default to "000"
            };
            return $"{disciplinePrefix}{formattedNumber}";
        }
        
        /// <summary>
        /// Set MEP Mark parameter on element
        /// </summary>
        private void SetMarkParameter(Element element, string markValue)
        {
            // Resolve parameter name case/space/underscore-insensitively
            var markParam = ResolveMarkParameter(element);

            File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\mepmark_debug.log", 
                $"[SET-DEBUG] Element {element.Id}: ResolvedParam='{markParam?.Definition?.Name}', Value='{markValue}', Doc.IsModifiable={element.Document.IsModifiable}\n");

            if (markParam != null && !markParam.IsReadOnly && markParam.StorageType == StorageType.String)
            {
                markParam.Set(markValue);
                // Verify write stuck
                try { element.Document.Regenerate(); } catch {}
                var readBack = markParam.AsString();
                if (!string.Equals(readBack, markValue, StringComparison.Ordinal))
                {
                    // Retry once
                    try { element.Document.Regenerate(); } catch {}
                    markParam.Set(markValue);
                    try { element.Document.Regenerate(); } catch {}
                    readBack = markParam.AsString();
                }
                if (string.Equals(readBack, markValue, StringComparison.Ordinal))
                {
                    File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\mepmark_debug.log", 
                        $"[SET-SUCCESS] Applied '{markValue}' to element {element.Id} using '{markParam.Definition?.Name}'\n");
                }
                else
                {
                    File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\mepmark_debug.log", 
                        $"[SET-VERIFY-FAIL] Element {element.Id}: attempted '{markValue}', read-back='{readBack ?? "<null>"}'\n");
                    throw new InvalidOperationException($"MEP Mark write did not persist on element {element.Id.IntegerValue}");
                }
            }
            else
            {
                var readonlyInfo = markParam != null ? $"Param='{markParam.Definition?.Name}', Readonly={markParam.IsReadOnly}, Type={markParam.StorageType}" : "Param=null";
                File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\mepmark_debug.log", 
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

