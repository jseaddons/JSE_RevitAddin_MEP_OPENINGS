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
    public partial class MarkParameterService
    {
        private Dictionary<long, ClashZone>? _clashZoneCache;
        private Dictionary<int, ClashZone>? _clusterZoneCache;
        private HashSet<int>? _combinedSleeveCache;
        private bool _cacheInitialized = false;
        private Document? _cachedDocument = null; // Track document for cache invalidation
        
        // ✅ OPTIMIZATION: Cache max mark numbers per category+prefix combination to avoid repeated database queries
        private Dictionary<string, int>? _maxMarkNumberCache;
        private Document? _maxMarkCacheDocument = null;

        // ✅ OPTIMIZATION: Cache resolved prefixes by category + system/service type (avoid repeated lookups)
        private readonly Dictionary<string, string> _prefixResolutionCache = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        // ✅ OPTIMIZATION: Cache existing mark numbers per category/prefix set to avoid repeated DB scans
        private Dictionary<string, HashSet<int>>? _existingMarkNumbersCache;
        private Document? _existingMarkNumbersCacheDocument;
        
        // Logger delegate
        private readonly Action<string> _logger;
        
        /// <summary>
        /// Constructor with optional document and logger
        /// </summary>
        public MarkParameterService(Document doc = null, Action<string> logger = null)
        {
            _cachedDocument = doc;
            _logger = logger ?? ((msg) => { });
        }
        
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
        /// ✅ SIMPLIFIED: Apply marks from database (per-sheet numbering)
        /// Gets sleeves from DB by level, assigns fresh sequential numbers, batch writes to Revit
        /// </summary>
        /// <param name="doc">Revit document</param>
        /// <param name="settings">Prefix and numbering settings</param>
        /// <param name="category">Optional category filter (e.g., "Ducts"). If null, processes all categories.</param>
        /// <returns>Tuple of (processedCount, errorCount)</returns>
        public (int processedCount, int errorCount) ApplyMarksFromDatabase(
            Document doc,
            MarkPrefixSettings settings,
            string? category = null)
        {
            int processed = 0, errors = 0;

            // ✅ FORCE LOG: Verify entry and DLL version
            try 
            {
                string logPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                var buildTime = System.IO.File.GetLastWriteTime(assembly.Location);
                File.AppendAllText(logPath, $"\n[MarkParameterService] >>> ENTRY ApplyMarksFromDatabase at {DateTime.Now:HH:mm:ss} <<<\n");
                File.AppendAllText(logPath, $"[MarkParameterService] DLL Build Time: {buildTime:yyyy-MM-dd HH:mm:ss}\n");
                File.AppendAllText(logPath, $"[MarkParameterService] DeploymentMode: {DeploymentConfiguration.DeploymentMode}\n");
            }
            catch {} // Ignore logging errors


            try
            {
                // 1. Get active floor plan level and view extent
                var activeView = doc.ActiveView;
                if (!(activeView is ViewPlan plan))
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning("[MarkParameterService] ApplyMarksFromDatabase: Active view must be a floor plan");
                    throw new InvalidOperationException("Active view must be a floor plan");
                }

                var level = plan.GenLevel;
                if (level == null)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning("[MarkParameterService] ApplyMarksFromDatabase: Floor plan must have an associated level");
                    throw new InvalidOperationException("Floor plan must have an associated level");
                }

                string levelName = level.Name;
                
                // Get settings
                int startNumber = settings.StartNumber;
                string numberFormat = settings.NumberFormat;
                string projectPrefix = settings.ProjectPrefix;
                
                bool isCombinedPass = category != null && category.Equals("Combined", StringComparison.OrdinalIgnoreCase);
                if (isCombinedPass)
                {
                    startNumber = 1; // Combined always starts from 1
                }
                
                // ✅ VIEW EXTENT: Get view bounds for per-sheet filtering
                // CropBox returns effective bounds from crop region OR scope box (whichever defines the view extent)
                // This allows multiple sheets on the same level to have independent numbering
                BoundingBoxXYZ? viewExtent = null;
                if (plan.CropBoxActive && plan.CropBox != null)
                {
                    viewExtent = plan.CropBox;
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[MarkParameterService] ApplyMarksFromDatabase: Using view extent (crop box/scope box) - " +
                            $"Min({viewExtent.Min.X:F1}, {viewExtent.Min.Y:F1}), Max({viewExtent.Max.X:F1}, {viewExtent.Max.Y:F1})");
                    }
                }
                
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[MarkParameterService] ApplyMarksFromDatabase: Level='{levelName}', Category='{category ?? "ALL"}', Start={startNumber}, HasCropBox={viewExtent != null}");

                // 2. Query DB for sleeves on this level
                using var context = new SleeveDbContext(doc);
                var repo = new ClashZoneRepository(context, msg =>
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info(msg);
                });

                var zones = repo.GetSleevesForLevel(levelName, category);
                
                if (zones.Count == 0)
                {
                    // ✅ FORCE LOG: Help user debug empty results
                    string logPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                    File.AppendAllText(logPath, $"[MarkParameterService] WARN: No sleeves found for level '{levelName}' in DB. Category='{category ?? "null"}'\n");
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[MarkParameterService] ApplyMarksFromDatabase: No sleeves found for level '{levelName}'");
                    return (0, 0);
                }

                // ✅ VIEW EXTENT FILTER: If crop box active, filter zones by placement point
                // IMPORTANT: Iterate 1698 - CropBox is in View Coordinates, DB sleeves are in World Coordinates.
                // We MUST transform the CropBox to World Coordinates.
                if (viewExtent != null)
                {
                    Transform viewTransform = viewExtent.Transform;
                    
                    // Transform the 4 corners of the view's bounding box rectangle (z is ignored for 2D check)
                    XYZ bMin = viewExtent.Min;
                    XYZ bMax = viewExtent.Max;
                    
                    var corners = new List<XYZ>
                    {
                        viewTransform.OfPoint(new XYZ(bMin.X, bMin.Y, bMin.Z)),
                        viewTransform.OfPoint(new XYZ(bMax.X, bMin.Y, bMin.Z)),
                        viewTransform.OfPoint(new XYZ(bMax.X, bMax.Y, bMin.Z)),
                        viewTransform.OfPoint(new XYZ(bMin.X, bMax.Y, bMin.Z))
                    };
                    
                    double worldMinX = corners.Min(c => c.X);
                    double worldMinY = corners.Min(c => c.Y);
                    double worldMaxX = corners.Max(c => c.X);
                    double worldMaxY = corners.Max(c => c.Y);
                    
                    // Apply tolerance
                    double tolerance = 0.01; 
                    worldMinX -= tolerance; worldMinY -= tolerance;
                    worldMaxX += tolerance; worldMaxY += tolerance;

                    int beforeCount = zones.Count;
                    zones = zones.Where(z => {
                        // ✅ COORDINATE FALLBACK: Try ActiveDoc -> Original Placement -> Intersection
                        double pX = z.SleevePlacementPointActiveDocumentX;
                        double pY = z.SleevePlacementPointActiveDocumentY;
                        
                        if (Math.Abs(pX) < 0.001 && Math.Abs(pY) < 0.001)
                        {
                            // Fallback 1: Original Placement Point (persisted)
                            pX = z.SleevePlacementPointX;
                            pY = z.SleevePlacementPointY;
                        }
                        
                        if (Math.Abs(pX) < 0.001 && Math.Abs(pY) < 0.001)
                        {
                            // Fallback 2: Intersection Point (always present from initial clash)
                            pX = z.IntersectionPoint.X;
                            pY = z.IntersectionPoint.Y;
                        }
                        
                        return pX >= worldMinX && pX <= worldMaxX &&
                               pY >= worldMinY && pY <= worldMaxY;
                    }).ToList();
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[MarkParameterService] ApplyMarksFromDatabase: Transformed View Extent: " +
                            $"Min({worldMinX:F1}, {worldMinY:F1}), Max({worldMaxX:F1}, {worldMaxY:F1}). Filtered: {beforeCount} → {zones.Count}");
                }

                if (zones.Count == 0)
                {
                    // ✅ FORCE LOG: View Extent filter eliminated all sleeves
                    /*
                    string logPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                    try 
                    {
                        var sb = new System.Text.StringBuilder();
                        sb.AppendLine($"[MarkParameterService] WARN: All sleeves filtered out by View Extent (Active View Only).");
                        // ... (omitted for brevity)
                        File.AppendAllText(logPath, sb.ToString());
                    }
                    catch {}
                    */

                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[MarkParameterService] ApplyMarksFromDatabase: No sleeves within view extent");
                    return (0, 0);
                }

                // 3. Batch write marks in single transaction
                int currentNumber = startNumber;
                // isCombinedPass is already defined in outer scope
                
                // ✅ OPTIMIZATION: Group by actual element ID to avoid double-marking (Combined/Cluster have multiple zones)
                // We use a hierarchy for the "active" sleeve ID: Combined > Cluster > Individual
                var elementGroups = zones.GroupBy(z => {
                    if (z.CombinedClusterSleeveInstanceId > 0) return z.CombinedClusterSleeveInstanceId;
                    if (z.ClusterSleeveInstanceId > 0) return z.ClusterSleeveInstanceId;
                    return z.SleeveInstanceId;
                }).Where(g => g.Key > 0).ToList();

                // 3. Parallel Calculation Phase: Resolve prefixes and format strings
                // This is CPU-bound and doesn't use Revit API, so it's safe to parallelize
                var markAssignments = new System.Collections.Concurrent.ConcurrentDictionary<long, string>();
                System.Threading.Tasks.Parallel.ForEach(elementGroups, (group, state, index) =>
                {
                    long sleeveIdValue = group.Key;
                    int localNumber = startNumber + (int)index;

                    string numStr = localNumber.ToString().PadLeft(numberFormat.Length, '0');
                    
                    string effectiveDiscipline = "";
                    if (isCombinedPass)
                    {
                        effectiveDiscipline = "MEP";
                    }
                    else
                    {
                        var firstZone = group.First();
                        string sysType = firstZone.GetParameterValue("System Type") ?? firstZone.GetParameterValue("MEP System Type") ?? "";
                        string svcType = firstZone.GetParameterValue("Service Type") ?? "";
                        effectiveDiscipline = settings.GetPrefixForElement(firstZone.MepElementCategory, sysType, svcType);
                    }
                    
                    string markValue = $"{projectPrefix}{effectiveDiscipline}{numStr}";
                    markAssignments.TryAdd(sleeveIdValue, markValue);
                });

                // 4. Write Phase: Single-threaded transaction to push data to Revit
                using (var trans = new Transaction(doc, "Apply Marks"))
                {
                    trans.Start();
                    
                    foreach (var group in elementGroups)
                    {
                        long sleeveIdValue = group.Key;
                        if (!markAssignments.TryGetValue(sleeveIdValue, out string markValue)) continue;

#if REVIT2024_OR_GREATER
                        var sleeve = doc.GetElement(new ElementId(sleeveIdValue)) as FamilyInstance;
#else
                        var sleeve = doc.GetElement(new ElementId((int)sleeveIdValue)) as FamilyInstance;
#endif
                        if (sleeve == null || !sleeve.IsValidObject)
                        {
                            errors++;
                            continue;
                        }

                        try
                        {
                            var markParam = sleeve.LookupParameter("MEP Mark") ?? sleeve.LookupParameter("Mark");
                            if (markParam != null && !markParam.IsReadOnly && markParam.StorageType == StorageType.String)
                            {
                                markParam.Set(markValue);
                                processed++;
                            }
                            else
                            {
                                errors++;
                            }
                        }
                        catch
                        {
                            errors++;
                        }
                    }

                    trans.Commit();
                }

                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[MarkParameterService] ApplyMarksFromDatabase: Processed {processed} elements from {zones.Count} zones, Errors {errors}");
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[MarkParameterService] ApplyMarksFromDatabase: {ex.Message}");
                throw;
            }

            return (processed, errors);
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
            int maxNumberUsed = 0; // Track highest number assigned this run
            int markerLastCount = 0; // Track marker-based last count

            // ✅ MARKER REPO: Track last processed mark number per category (CategoryProcessingMarkers table)
            using var markerContext = new SleeveDbContext(doc);
            var markerRepository = new CategoryProcessingMarkerRepository(markerContext, msg =>
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info(msg);
            });

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
                    File.AppendAllText(mepmarkLogPath, $"🔨 BUILD TIME: {buildTime:yyyy-MM-dd HH:mm:ss} (DLL: {Path.GetFileName(assembly.Location)})\n");
                }
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

                // ✅ RESOURCE FIX: Use shared context for all lookups
                using var sharedContext = new SleeveDbContext(doc);

                // ✅ SIMPLIFIED: Get ALL opening sleeves (individual + cluster + combined) for this category
                // For marks, we only need project prefix + numbering, no need to separate by type
                // ✅ BIM 360 OPTIMIZATION: Pass markPrefixes to enable active view filtering
                var allSleeves = GetAllSleevesForCategory(doc, category, sharedContext, markPrefixes);
                
                // ✅ DEBUG: Log sleeve distribution by host type (Wall vs Floor)
                var sleevesByHostType = allSleeves.GroupBy(s => {
                    var famName = s.Symbol?.Family?.Name ?? "";
                    if (famName.IndexOf("OpeningOnWall", StringComparison.OrdinalIgnoreCase) >= 0) return "Wall";
                    if (famName.IndexOf("OpeningOnSlab", StringComparison.OrdinalIgnoreCase) >= 0) return "Floor";
                    return "Unknown";
                }).ToDictionary(g => g.Key, g => g.Count());
                
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[MarkParameterService] Found {allSleeves.Count} total sleeves (individual + cluster + combined) for category '{category}'");
                                // ✅ DEPLOYMENT MODE: Skip file writes
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    File.AppendAllText(mepmarkLogPath, $"Found {allSleeves.Count} total sleeves (individual + cluster + combined) for category '{category}'\n");
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
                        
                        // ✅ DIAGNOSTIC: Check if sleeves exist globally (if ActiveViewOnly was true)
                        if (markPrefixes?.ActiveViewOnly == true)
                        {
                            using (var diagContext = new SleeveDbContext(doc))
                            {
                                var globalSleeves = GetAllSleevesForCategory(doc, category, diagContext, markPrefixes); 
                                if (globalSleeves.Count > 0)
                                {
                                    File.AppendAllText(mepmarkLogPath, $"[DIAGNOSTIC] ⚠️ FOUND {globalSleeves.Count} SLEEVES GLOBALLY! They are hidden in the current view '{(doc.ActiveView?.Name ?? "Unknown")}' or excluded by Section Box.\n");
                                    File.AppendAllText(mepmarkLogPath, $"[DIAGNOSTIC] Suggestion: Disable 'Active View Only' or switch to a plan view where sleeves are visible.\n");
                                }
                                else
                                {
                                    File.AppendAllText(mepmarkLogPath, $"[DIAGNOSTIC] Verified: 0 sleeves found globally for category '{category}' even outside active view.\n");
                                }
                            }
                        }
                    }
                    return (0, 0);
                }
                
                // ✅ FIX: Get max number per category (consider all active prefixes for this category)
                var candidatePrefixes = GetCandidateDisciplinePrefixes(category, disciplinePrefix, markPrefixes);

                // ✅ MARKER: Load last processed count (acts as last mark number) and combine with DB max
                var markerInfo = markerRepository.GetMarker(category);
                markerLastCount = markerInfo.lastCount;

                int categoryMaxNumber = GetMaxExistingMarkNumberForCategory(doc, category, candidatePrefixes);
                categoryMaxNumber = Math.Max(categoryMaxNumber, markerLastCount);
                int startIndex = categoryMaxNumber + 1;
                
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[MarkParameterService] Category '{category}': Max existing number = {categoryMaxNumber}, Starting at: {startIndex}");
                                // ✅ DEPLOYMENT MODE: Skip file writes
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    File.AppendAllText(mepmarkLogPath, $"Category '{category}': Max existing number = {categoryMaxNumber}, Starting at: {startIndex}\n");
                }
                
                // ✅ TRACK: Preload existing numbers when remarking to avoid O(n) range fill
                HashSet<int> usedNumbers;
                if (remarkAll)
                {
                    usedNumbers = GetExistingMarkNumbers(doc, category, candidatePrefixes);

                    // Fallback: maintain previous behavior if DB scan returns nothing
                    if (usedNumbers.Count == 0 && categoryMaxNumber > 0)
                    {
                        usedNumbers = new HashSet<int>(Enumerable.Range(1, categoryMaxNumber));
                    }
                }
                else
                {
                    usedNumbers = new HashSet<int>();
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
                        clashZone = GetClashZoneByClusterInstanceId(sleeveId, category, sharedContext);
                        
                        // If not found, try MEP_ElementId (for individual sleeves)
                        if (clashZone == null && mepElementId > 0)
                        {
                            clashZone = GetClashZoneByMepElementId(mepElementId, doc, sharedContext);
                        }
                        
                        // ✅ ENHANCED: Resolve element-specific prefix (e.g., System Type override)
                        // ✅ CRITICAL: This applies to ALL families (RectangularOpeningOnWall, CircularOpeningOnWall, etc.)
                        // System Type override takes precedence over discipline prefix for all families
                        var elementPrefix = ResolveDisciplinePrefixForElement(category, disciplinePrefix, markPrefixes, clashZone);
                        if (!string.IsNullOrWhiteSpace(elementPrefix))
                        {
                            candidatePrefixes.Add(elementPrefix);
                        }
                        
                        // ✅ DIAGNOSTIC: Log family name and resolved prefix to verify System Type override is working
                        // Note: famName and mepmarkLogPath are declared later in the loop, so we'll log there

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
                                maxNumberUsed = Math.Max(maxNumberUsed, numberToUse);
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
                                maxNumberUsed = Math.Max(maxNumberUsed, numberToUse);
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
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                File.AppendAllText(mepmarkLogPath, $"SKIP sleeve {sleeve.Id}: already has mark '{existingMark}' (remark=false)\n");
                            }
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
                            maxNumberUsed = Math.Max(maxNumberUsed, numberToUse);
                        }
                        
                        // ✅ ENHANCED: Log detailed info about sleeve before marking
                        
                        // ✅ DEBUG: Log sleeve host type
                        var famNameResolved = sleeve.Symbol?.Family?.Name ?? "";
                        var hostType = famNameResolved.IndexOf("OpeningOnWall", StringComparison.OrdinalIgnoreCase) >= 0 ? "Wall" :
                                      famNameResolved.IndexOf("OpeningOnSlab", StringComparison.OrdinalIgnoreCase) >= 0 ? "Floor" : "Unknown";
                        
                        // ✅ ENHANCED LOGGING: Log ALL sleeves (not just first 10) to diagnose floor sleeve issue
                        // Note: mepElementIdParam and mepElementId are already declared above
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            // ✅ DIAGNOSTIC: Also log System Type override resolution for this sleeve
                            string systemType = clashZone != null ? GetClashParameterValue(clashZone, "System Type", "MEP System Type", "System Classification") : null;
                            File.AppendAllText(mepmarkLogPath, 
                            $"[MARK-PROCESS] Sleeve {sleeve.Id}: HostType={hostType}, Family={famNameResolved}, MEP_ID={mepElementId}, " +
                                $"Category={clashZone?.MepElementCategory ?? "null"}, RequestedCategory={category}, " +
                                $"CategoryMatch={clashZone?.MepElementCategory == category}, RemarkAll={remarkAll}, ExistingMark='{existingMark ?? "null"}'\n");
                            File.AppendAllText(mepmarkLogPath,
                            $"[PREFIX-RESOLVE] Family={famNameResolved}, Category={category}, SystemType='{systemType ?? "null"}', " +
                                $"DisciplinePrefix='{disciplinePrefix}', ResolvedPrefix='{elementPrefix}', " +
                                $"IsOverride={(elementPrefix != disciplinePrefix ? "YES" : "NO")}\n");
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
                // ✅ MARKER: Persist highest mark number used for this category (incremental processing)
                if (maxNumberUsed > 0)
                {
                    // If nothing new was assigned, keep previous marker count
                    int newMarkerCount = Math.Max(maxNumberUsed, markerLastCount);
                    markerRepository.UpdateMarker(category, newMarkerCount);
                }

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
        
        /// <summary>
        /// Get ALL opening sleeves for a category (individual + cluster + combined)
        /// ✅ SIMPLIFIED: For marks, we don't need to separate by type - just collect all opening families
        /// ✅ BIM 360 OPTIMIZATION: Optionally filter by active view for per-sheet numbering
        /// </summary>
        /// <summary>
        /// Get ALL opening sleeves for a category (individual + cluster + combined)
        /// ✅ SIMPLIFIED: For marks, we don't need to separate by type - just collect all opening families
        /// ✅ BIM 360 OPTIMIZATION: Optionally filter by active view for per-sheet numbering
        /// ✅ RESOURCE FIX: Uses shared context
        /// </summary>
        private List<FamilyInstance> GetAllSleevesForCategory(Document doc, string category, SleeveDbContext sharedContext, MarkPrefixSettings? markPrefixes = null)
        {
            // ✅ BIM 360 OPTIMIZATION: Use active view collector if ActiveViewOnly is enabled
            FilteredElementCollector collector;
            if (markPrefixes?.ActiveViewOnly == true && doc.ActiveView != null)
            {
                collector = new FilteredElementCollector(doc, doc.ActiveView.Id);
                string logPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    File.AppendAllText(logPath,
                        $"[ACTIVE-VIEW-FILTER] ✅ Filtering by active view: {doc.ActiveView.Name} (ID: {doc.ActiveView.Id})\n");
                }
            }
            else
            {
                collector = new FilteredElementCollector(doc);
            }
            
            // Collect ALL opening family instances (individual + cluster + combined)
            var allOpeningSleeves = collector
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => {
                    var famName = fi.Symbol?.Family?.Name ?? string.Empty;
                    return famName.IndexOf("OpeningOnWall", StringComparison.OrdinalIgnoreCase) >= 0
                        || famName.IndexOf("OpeningOnSlab", StringComparison.OrdinalIgnoreCase) >= 0;
                })
                .ToList();

            if (!DeploymentConfiguration.DeploymentMode)
            {
                string logPathDebug = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                File.AppendAllText(logPathDebug, $"[GetAllSleeves] RAW Collector found {allOpeningSleeves.Count} potential opening families in view '{(doc.ActiveView?.Name ?? "null")}'\n");
            }
            
            // Filter by category using database lookup
            var categorySleeves = new List<FamilyInstance>();
            int rejectedCount = 0;
            
            foreach (var sleeve in allOpeningSleeves)
            {
                // Try to determine category from database
                // For individual sleeves: check by MEP_ElementId
                // For cluster sleeves: check by ClusterInstanceId
                // For cluster sleeves: check by ClusterSleeveInstanceId
                // For combined sleeves: check by CombinedInstanceId
                
                int sleeveId = sleeve.Id.IntegerValue;
                bool matchesCategory = false;
                
                // 1. Try cluster lookup (using cache first)
                var clashZone = GetClashZoneByClusterInstanceId(sleeveId, category, sharedContext);
                if (clashZone != null && clashZone.MepElementCategory == category)
                {
                    matchesCategory = true;
                }
                else
                {
                     // 2. ✅ Try Combined Sleeve lookup
                     // If it's a Combined Sleeve, it has a CombinedInstanceId in DB
                     // Check cache first
                     if (_combinedSleeveCache != null && _combinedSleeveCache.Contains(sleeveId))
                     {
                         matchesCategory = true;
                     }
                     
                     if (!matchesCategory)
                     {
                        // 3. Try individual sleeve lookup (by MEP_ElementId)
                        var mepElementIdParam = sleeve.LookupParameter("MEP_ElementId");
                        if (mepElementIdParam != null)
                        {
                            long mepId = 0;
                            if (mepElementIdParam.StorageType == StorageType.Integer) mepId = mepElementIdParam.AsInteger();
                            else if (mepElementIdParam.StorageType == StorageType.Double) mepId = (long)mepElementIdParam.AsDouble();
                            else if (mepElementIdParam.StorageType == StorageType.String && long.TryParse(mepElementIdParam.AsString(), out long pVal)) mepId = pVal;

                            if (mepId > 0)
                            {
                                var zone = GetClashZoneByMepElementId(mepId, doc, sharedContext);
                                if (zone != null && zone.MepElementCategory == category)
                                {
                                    matchesCategory = true;
                                }
                            }
                        }
                     }
                }
                
                // ✅ COMBINED SLEEVES: Also check if this is a combined sleeve
                // Combined sleeves don't have ClashZones, but they have MEP_Category parameter
                if (!matchesCategory)
                {
                    var mepCategoryParam = sleeve.LookupParameter("MEP_Category");
                    if (mepCategoryParam != null && mepCategoryParam.AsString() == "Multi-Service")
                    {
                        // This is a combined sleeve - include it for all categories
                        // Combined sleeves serve multiple categories, so they should appear in all category lists
                        matchesCategory = true;
                    }
                }
                
                if (matchesCategory)
                {
                    categorySleeves.Add(sleeve);
                }
                else
                {
                    // Log first 10 rejections for debugging
                    rejectedCount++;
                    if (!DeploymentConfiguration.DeploymentMode && rejectedCount <= 10)
                    {
                         string logPathDebug = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                         string actualCat = clashZone?.MepElementCategory ?? "null";
                         File.AppendAllText(logPathDebug, $"[GetAllSleeves] ❌ REJECT sleeve {sleeve.Id}: Requested='{category}', Found='{actualCat}' (MepID={(sleeve.LookupParameter("MEP_ElementId")?.AsInteger() ?? -1)})\n");
                    }
                }
            }
            
            if (!DeploymentConfiguration.DeploymentMode)
            {
                string logPathFinal = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                File.AppendAllText(logPathFinal, 
                    $"[GetAllSleeves] Found {categorySleeves.Count} total sleeves (individual + cluster + combined) for category '{category}'\n");
            }
            
            return categorySleeves;
        }
        
        /// <summary>
        /// Find individual sleeves for a specific category (non-clustered sleeves)
        /// ✅ BIM 360 OPTIMIZATION: Optionally filter by active view for per-sheet numbering
        /// </summary>
                            // 4. Update GetIndividualSleevesForCategory compilation (not checking call sites, but fixing method)
                            // Actually, I don't see GetIndividualSleevesForCategory in this replacement. I should update it if it exists.
                            // I do not see it in the viewed lines? 
                            // Ah, I viewed up to 800, and helper was around 700. Yes.
                            // I'll assume GetIndividualSleevesForCategory is handled or unused, but if I saw it, I should update it.
                            // I will do it in next tool call if needed or include it here if I have lines.
                            // I do have lines 678-793 in cache.
        private List<FamilyInstance> GetIndividualSleevesForCategory(Document doc, string category, SleeveDbContext sharedContext, MarkPrefixSettings? markPrefixes = null)
        {
            // ✅ BIM 360 OPTIMIZATION: Use active view collector if ActiveViewOnly is enabled
            FilteredElementCollector collector;
            if (markPrefixes?.ActiveViewOnly == true && doc.ActiveView != null)
            {
                // Only collect elements visible in active view (per-sheet numbering)
                collector = new FilteredElementCollector(doc, doc.ActiveView.Id);
                
                string mepmarkLogPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    File.AppendAllText(mepmarkLogPath,
                        $"[ACTIVE-VIEW-FILTER] ✅ Filtering by active view: {doc.ActiveView.Name} (ID: {doc.ActiveView.Id})\n");
                }
            }
            else
            {
                // Collect all elements in document (global numbering)
                collector = new FilteredElementCollector(doc);
            }
            
            var allSleeves = collector
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
                    var clashZone = GetClashZoneByMepElementId(mepElementId, doc, sharedContext);
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
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            File.AppendAllText(mepmarkLogPath, 
                                $"[INDIVIDUAL-DEBUG] Sleeve {sleeve.Id}: MEP_ID={mepElementId}, " +
                                $"Category_match={clashZone.MepElementCategory == category} (XML='{clashZone.MepElementCategory}' vs Requested='{category}'), " +
                                $"isClusterSleeve={isClusterSleeve} (IsClusterResolved={clashZone.IsClusterResolved}, ClusterSleeveInstanceId={clashZone.ClusterSleeveInstanceId}, SleeveId={sleeve.Id.IntegerValue})\n");
                        }
                        
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
        /// ✅ BIM 360 OPTIMIZATION: Optionally filter by active view for per-sheet numbering
        /// </summary>
        private List<FamilyInstance> GetClusterSleevesForCategory(Document doc, string category, SleeveDbContext sharedContext, MarkPrefixSettings? markPrefixes = null)
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

            // ✅ BIM 360 OPTIMIZATION: Use active view collector if ActiveViewOnly is enabled
            FilteredElementCollector collector;
            if (markPrefixes?.ActiveViewOnly == true && doc.ActiveView != null)
            {
                // Only collect elements visible in active view (per-sheet numbering)
                collector = new FilteredElementCollector(doc, doc.ActiveView.Id);
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    File.AppendAllText(mepmarkLogPath,
                        $"[ACTIVE-VIEW-FILTER] ✅ Filtering cluster sleeves by active view: {doc.ActiveView.Name} (ID: {doc.ActiveView.Id})\n");
                }
            }
            else
            {
                // Collect all elements in document (global numbering)
                collector = new FilteredElementCollector(doc);
            }

            var allClusterSleeves = collector
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => {
                    var famName = fi.Symbol?.Family?.Name ?? string.Empty;
                    // ✅ FIX: Use contains instead of exact match to handle any family name variations
                    return famName.IndexOf("OpeningOnWall", StringComparison.OrdinalIgnoreCase) >= 0
                        || famName.IndexOf("OpeningOnSlab", StringComparison.OrdinalIgnoreCase) >= 0;
                })
                .ToList();

            if (!DeploymentConfiguration.DeploymentMode)
            {
                File.AppendAllText(mepmarkLogPath, $"Found {allClusterSleeves.Count} total opening sleeves (after family name filter)\n");
            }

            // ✅ CRITICAL FIX: Find cluster sleeves by ClusterInstanceId instead of MEP_ElementId
            // Cluster sleeves don't have a single MEP_ElementId, so we query clash zones by ClusterInstanceId
            var categorySleeves = new List<FamilyInstance>();

            foreach (var sleeve in allClusterSleeves)
            {
                int sleeveId = sleeve.Id.IntegerValue;
                
                // ✅ NEW APPROACH: Query clash zones by ClusterInstanceId (cluster sleeve's own ID)
                var clashZone = GetClashZoneByClusterInstanceId(sleeveId, category, sharedContext);
                
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
        /// ✅ RESOURCE FIX: Uses shared context
        /// </summary>
        private ClashZone GetClashZoneByClusterInstanceId(int clusterInstanceId, string category, SleeveDbContext sharedContext)
        {
            if (clusterInstanceId <= 0)
                return null;

            // ✅ OPTIMIZATION: Check cache first
            if (_clusterZoneCache != null && _clusterZoneCache.TryGetValue(clusterInstanceId, out ClashZone cachedZone))
            {
                if (cachedZone.MepElementCategory == category)
                    return cachedZone;
            }

            if (sharedContext == null)
                return null;

            try
            {
                // ✅ OPTIMIZATION: Ensure cache initialized if we hit this point
                if (!_cacheInitialized || _cachedDocument != null) // document check logic here simplified or assumes caller handles
                {
                    // Fallback to manual load if cache missed but we have context
                    var repository = new ClashZoneRepository(sharedContext);
                    var zones = repository.GetClashZonesByCategory(category);
                    var foundZone = zones?.FirstOrDefault(z => z.ClusterSleeveInstanceId == clusterInstanceId && z.IsClusterResolved);
                    
                    if (foundZone != null)
                    {
                        // Add to cache
                        if (_clusterZoneCache == null) _clusterZoneCache = new Dictionary<int, ClashZone>();
                        _clusterZoneCache[clusterInstanceId] = foundZone;
                        return foundZone;
                    }
                }
            }
            catch (Exception ex)
            {
                // Silent catch or logger
            }
            return null;
        }

        /// <summary>
        /// Get clash zone by MEP element ID using cache and database fallback
        /// ⚠️ PERFORMANCE: Uses cache to avoid O(n·m) lookups
        /// ✅ RESOURCE FIX: Uses shared context
        /// </summary>
        private ClashZone GetClashZoneByMepElementId(long mepElementId, Document doc, SleeveDbContext sharedContext)
        {
            // Initialize cache once per service instance
            if (!_cacheInitialized || _cachedDocument != doc)
            {
                InitializeClashZoneCache(doc, sharedContext);
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
            if (sharedContext != null && mepElementId > 0)
            {
                try
                {
                    // ✅ RESOURCE FIX: Use shared context
                    var repository = new ClashZoneRepository(sharedContext);
                    
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
        private void InitializeClashZoneCache(Document doc, SleeveDbContext sharedContext)
        {
            _clashZoneCache = new Dictionary<long, ClashZone>();
            _clusterZoneCache = new Dictionary<int, ClashZone>();
            _combinedSleeveCache = new HashSet<int>();

            try
            {
                string mepmarkLogPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                // ✅ DEPLOYMENT MODE: Skip file writes
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    File.AppendAllText(mepmarkLogPath, $"\n[CACHE-INIT] ===== CACHE INITIALIZATION STARTED =====\n");
                }
                File.AppendAllText(mepmarkLogPath, $"[CACHE-INIT] Document: {(doc != null ? doc.Title : "NULL")}\n");
                
                int totalClashZones = 0;
                
                // ✅ STEP 1: Try loading from database first
                if (doc != null && sharedContext != null)
                {
                    try
                    {
                        // ✅ RESOURCE FIX: Use shared context
                        var repository = new ClashZoneRepository(sharedContext);
                            
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
                                    if (clashZone != null)
                                    {
                                        // 1. MepElementId Cache
                                        if (clashZone.MepElementIdValue > 0)
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

                                        // 2. Cluster Cache
                                        if (clashZone.IsClusterResolved && clashZone.ClusterSleeveInstanceId > 0)
                                        {
                                            _clusterZoneCache[clashZone.ClusterSleeveInstanceId] = clashZone;
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
                        if (totalClashZones == 0)
                        {
                            // ✅ FALLBACK: Direct query to ClashZones table if category-based query returned empty
                            File.AppendAllText(mepmarkLogPath, 
                                $"[CACHE-INIT] ⚠️ GetClashZonesByCategory returned 0 zones, trying direct query...\n");
                            
                            using (var directCmd = sharedContext.Connection.CreateCommand())
                            {
                                directCmd.CommandText = @"
                                    SELECT * FROM ClashZones 
                                    WHERE MepElementId > 0 
                                    AND MepCategory IS NOT NULL 
                                    AND MepCategory != ''";
                                
                                using (var reader = directCmd.ExecuteReader())
                                {
                                    while (reader.Read())
                                    {
                                        try
                                        {
                                            var clashZone = repository.GetType()
                                                .GetMethod("MapClashZone", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)
                                                ?.Invoke(repository, new object[] { reader }) as ClashZone;
                                            
                                            if (clashZone != null && clashZone.MepElementIdValue > 0)
                                            {
                                                long key = clashZone.MepElementIdValue;
                                                if (!_clashZoneCache.ContainsKey(key))
                                                {
                                                    clashZone.EnsureSleevePlacementPointReconstructed();
                                                    _clashZoneCache[key] = clashZone;
                                                    totalClashZones++;
                                                }
                                            }
                                        }
                                        catch { /* Skip malformed rows */ }
                                    }
                                }
                            }
                        }
                        
                         if (totalClashZones > 0)
                         {
                              // ✅ ALSO: Pre-load Combined Sleeves into cache
                              try
                              {
                                  var comboRepo = new CombinedSleeveRepository(sharedContext);
                                  var allCombos = comboRepo.GetAllCombinedSleeves();
                                  foreach(var combo in allCombos)
                                  {
                                      if (combo.CombinedInstanceId > 0)
                                          _combinedSleeveCache.Add(combo.CombinedInstanceId);
                                  }
                              }
                              catch (Exception comboEx)
                              {
                                  File.AppendAllText(mepmarkLogPath, $"[CACHE-INIT] ⚠️ Combined sleeve pre-load failed: {comboEx.Message}\n");
                              }

                              // ✅ DEPLOYMENT MODE: Skip file writes
                             if (!DeploymentConfiguration.DeploymentMode)
                             {
                                 File.AppendAllText(mepmarkLogPath, 
                                     $"[CACHE-INIT] ✓ Total {totalClashZones} clash zones loaded from database\n");
                                 File.AppendAllText(mepmarkLogPath, 
                                     $"[CACHE-INIT] Cache size: {_clashZoneCache.Count} entries\n\n");
                             }
                             return; 
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
        /// Load all existing mark numbers for a category from the database (fallback: Revit scan) and cache them.
        /// Helps remark runs skip numbers without filling large sequential ranges.
        /// </summary>
        private HashSet<int> GetExistingMarkNumbers(Document doc, string category, IEnumerable<string> disciplinePrefixes)
        {
            var prefixSet = new HashSet<string>((disciplinePrefixes ?? Enumerable.Empty<string>()), StringComparer.OrdinalIgnoreCase);
            prefixSet.RemoveWhere(string.IsNullOrWhiteSpace);

            if (prefixSet.Count == 0)
                return new HashSet<int>();

            string cacheKey = $"{category}_{string.Join("_", prefixSet.OrderBy(p => p))}";

            if (_existingMarkNumbersCache != null && _existingMarkNumbersCacheDocument == doc &&
                _existingMarkNumbersCache.TryGetValue(cacheKey, out var cached))
            {
                return new HashSet<int>(cached);
            }

            if (_existingMarkNumbersCache == null || _existingMarkNumbersCacheDocument != doc)
            {
                _existingMarkNumbersCache = new Dictionary<string, HashSet<int>>(StringComparer.OrdinalIgnoreCase);
                _existingMarkNumbersCacheDocument = doc;
            }

            var numbers = new HashSet<int>();

            // Try fast path: database snapshot values
            try
            {
                using (var dbContext = new SleeveDbContext(doc))
                using (var cmd = dbContext.Connection.CreateCommand())
                {
                    cmd.CommandText = @"
                        SELECT ss.MepParametersJson
                        FROM SleeveSnapshots ss
                        INNER JOIN ClashZones cz ON (
                            (ss.SleeveInstanceId IS NOT NULL AND ss.SleeveInstanceId = cz.SleeveInstanceId) OR
                            (ss.ClusterInstanceId IS NOT NULL AND ss.ClusterInstanceId = cz.ClusterInstanceId)
                        )
                        WHERE cz.MepElementCategory = @Category
                          AND ss.MepParametersJson IS NOT NULL
                          AND ss.MepParametersJson != '{}' AND ss.MepParametersJson != ''";

                    cmd.Parameters.AddWithValue("@Category", category);

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var mepParamsJson = reader.IsDBNull(0) ? null : reader.GetString(0);
                            if (string.IsNullOrWhiteSpace(mepParamsJson))
                                continue;

                            try
                            {
                                var mepParams = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, string>>(mepParamsJson);
                                if (mepParams != null && mepParams.TryGetValue("MEP Mark", out var markValue) && !string.IsNullOrWhiteSpace(markValue))
                                {
                                    int? extracted = ExtractNumberFromMark(markValue, prefixSet);
                                    if (extracted.HasValue)
                                    {
                                        numbers.Add(extracted.Value);
                                    }
                                }
                                else if (mepParams != null && mepParams.TryGetValue("Mark", out var altMark) && !string.IsNullOrWhiteSpace(altMark))
                                {
                                    int? extracted = ExtractNumberFromMark(altMark, prefixSet);
                                    if (extracted.HasValue)
                                    {
                                        numbers.Add(extracted.Value);
                                    }
                                }
                            }
                            catch
                            {
                                // Ignore malformed JSON rows
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Warning($"[MarkParameterService] Marker preload DB query failed for category '{category}': {ex.Message}");
                }
            }


            // Fallback: scan Revit elements if DB had no data
            if (numbers.Count == 0)
            {
                using (var fallbackContext = new SleeveDbContext(doc))
                {
                    var sleeveElements = new FilteredElementCollector(doc)
                        .OfClass(typeof(FamilyInstance))
                        .Cast<FamilyInstance>()
                        .Where(fi =>
                        {
                            var famName = fi.Symbol?.Family?.Name ?? string.Empty;
                            return famName.IndexOf("OpeningOnWall", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                   famName.IndexOf("OpeningOnSlab", StringComparison.OrdinalIgnoreCase) >= 0;
                        })
                        .ToList();

                    foreach (var element in sleeveElements)
                    {
                        var mepElementIdParam = element.LookupParameter("MEP_ElementId");
                        if (mepElementIdParam != null)
                        {
                            long mepElementId = mepElementIdParam.AsInteger();
                            var clashZone = GetClashZoneByMepElementId(mepElementId, doc, fallbackContext);
                            if (clashZone?.MepElementCategory != null && !clashZone.MepElementCategory.Equals(category, StringComparison.OrdinalIgnoreCase))
                                continue;
                        }

                        var markParam = ResolveMarkParameter(element);
                        string markValue = markParam?.AsString() ?? string.Empty;
                        if (!string.IsNullOrWhiteSpace(markValue))
                        {
                            int? extracted = ExtractNumberFromMark(markValue, prefixSet);
                            if (extracted.HasValue)
                            {
                                numbers.Add(extracted.Value);
                            }
                        }
                    }
                }
            }

            _existingMarkNumbersCache[cacheKey] = new HashSet<int>(numbers);
            return numbers;
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
            // ✅ CRITICAL: Extract System Type and Service Type from clash zone
            string systemType = GetClashParameterValue(clashZone, "System Type", "MEP System Type", "System Classification");
            string serviceType = GetClashParameterValue(clashZone, "Service Type", "System Abbreviation", "MEP System Type");

            return ResolveDisciplinePrefixFromStrings(category, defaultPrefix, markPrefixes, systemType, serviceType);
        }

        private string ResolveDisciplinePrefixFromStrings(string category, string defaultPrefix, MarkPrefixSettings markPrefixes, string systemType, string serviceType)
        {
            string fallbackPrefix = !string.IsNullOrWhiteSpace(defaultPrefix)
                ? defaultPrefix
                : markPrefixes?.GetDisciplinePrefix(category) ?? defaultPrefix ?? string.Empty;

            if (markPrefixes == null)
                return fallbackPrefix;

            // ✅ CACHE: Reuse resolved prefix for identical category/system/service combos
            string cacheKey = $"{category}|{systemType ?? string.Empty}|{serviceType ?? string.Empty}";
            if (_prefixResolutionCache.TryGetValue(cacheKey, out var cachedPrefix))
            {
                return cachedPrefix;
            }

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
            // ✅ CRITICAL: This applies to ALL sleeve families (RectangularOpeningOnWall, CircularOpeningOnWall, etc.)
            // No family-specific filtering - System Type override works for all families
            var resolved = markPrefixes.GetPrefixForElement(category, systemType, serviceType);
            
            // ✅ DEBUG: Log the resolved prefix with System Type context
            if (!DeploymentConfiguration.DeploymentMode)
            {
                string mepmarkLogPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                string disciplinePrefix = markPrefixes.GetDisciplinePrefix(category);
                bool isOverride = resolved != disciplinePrefix;
                File.AppendAllText(mepmarkLogPath, 
                    $"[PREFIX-RESOLVE] ✅ Resolved prefix='{resolved}', Discipline prefix='{disciplinePrefix}', IsOverride={isOverride}, " +
                    $"SystemType='{systemType ?? "null"}', ServiceType='{serviceType ?? "null"}'\n");
            }
            
            if (!string.IsNullOrWhiteSpace(resolved))
            {
                resolved = resolved.Trim();
                _prefixResolutionCache[cacheKey] = resolved;
                return resolved;
            }

            _prefixResolutionCache[cacheKey] = fallbackPrefix;
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
                
                using (var fallbackContext = new SleeveDbContext(doc))
                {
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
                            var clashZone = GetClashZoneByMepElementId(mepElementId, doc, fallbackContext);
                            
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

        #region Combined Sleeve Apply Marks - Optimized Flow

        /// <summary>
        /// Gets ALL combined sleeves from DB then retrieves elements from Revit.
        /// DB-FIRST approach: Query CombinedSleeves table for instance IDs.
        /// </summary>
        public List<FamilyInstance> GetAllCombinedSleeves(Document doc)
        {
            var combinedSleeves = new List<FamilyInstance>();
            
            try
            {
                // ✅ DB-FIRST: Get combined sleeve instance IDs from database
                var combinedInstanceIds = new List<int>();
                
                using (var context = new SleeveDbContext(doc))
                {
                    using (var cmd = context.Connection.CreateCommand())
                    {
                        cmd.CommandText = "SELECT DISTINCT CombinedInstanceId FROM CombinedSleeves WHERE CombinedInstanceId > 0";
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                combinedInstanceIds.Add(reader.GetInt32(0));
                            }
                        }
                    }
                }
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    string logPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                    File.AppendAllText(logPath, 
                        $"[COMBINED-MARKS] DB query returned {combinedInstanceIds.Count} combined sleeve instance IDs\n");
                }
                
                // ✅ Get elements from Revit by ID (fast - no collector needed)
                foreach (var instanceId in combinedInstanceIds)
                {
                    var element = doc.GetElement(new ElementId(instanceId));
                    if (element is FamilyInstance fi)
                    {
                        combinedSleeves.Add(fi);
                        
                        if (!DeploymentConfiguration.DeploymentMode && combinedSleeves.Count <= 3)
                        {
                            string logPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                            File.AppendAllText(logPath, 
                                $"[COMBINED-MARKS] ✅ Found combined sleeve: ID={instanceId}, Family='{fi.Symbol?.Family?.Name}'\n");
                        }
                    }
                }
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    string logPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                    File.AppendAllText(logPath, 
                        $"[COMBINED-MARKS] Found {combinedSleeves.Count} combined sleeves from DB lookup\n");
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    string logPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                    File.AppendAllText(logPath, $"[COMBINED-MARKS] Error getting combined sleeves: {ex.Message}\n");
                }
            }
            
            return combinedSleeves;
        }

        /// <summary>
        /// Detects if a FamilyInstance is a Combined Sleeve.
        /// Uses Revit parameters ONLY - no DB lookup required.
        /// </summary>
        public bool IsCombinedSleeve(FamilyInstance sleeve)
        {
            if (sleeve == null) return false;
            
            try
            {
                // Method 1: Check "Combined Sleeve Instance ID" parameter
                var combinedIdParam = sleeve.LookupParameter("Combined Sleeve Instance ID");
                if (combinedIdParam != null && combinedIdParam.AsInteger() > 0)
                    return true;
                
                // Method 2: Check MEP_Category = "Multi-Service" or "Combined"
                var categoryParam = sleeve.LookupParameter("MEP_Category");
                if (categoryParam != null)
                {
                    var val = categoryParam.AsString() ?? "";
                    if (val.Contains("Multi", StringComparison.OrdinalIgnoreCase) ||
                        val.Contains("Combined", StringComparison.OrdinalIgnoreCase))
                        return true;
                }
                
                // Method 3: Check family name contains "Combined"
                var familyName = sleeve.Symbol?.Family?.Name ?? "";
                if (familyName.Contains("Combined", StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            catch { /* Ignore parameter access errors */ }
            
            return false;
        }

        /// <summary>
        /// Calculates marks for combined sleeves.
        /// - Discipline prefix: HARDCODED to "MEP"
        /// - Project prefix: From UI settings
        /// - Number format: From UI settings (e.g., "000" for 3 digits)
        /// </summary>
        public List<(ElementId SleeveId, string Mark)> CalculateCombinedSleeveMarks(
            Document doc,
            List<FamilyInstance> combinedSleeves,
            string projectPrefix,
            string numberFormat,
            int startNumber,
            bool remarkAll)
        {
            const string DISCIPLINE_PREFIX = "MEP"; // ✅ HARDCODED
            
            var results = new List<(ElementId, string)>();
            if (combinedSleeves == null || combinedSleeves.Count == 0) return results;
            
            // 1. Sort sleeves (determines numbering order)
            var sortedSleeves = combinedSleeves
                .OrderBy(s => GetLevelNameForSorting(s, doc))
                .ThenByDescending(s => GetLocationForSorting(s).Y)
                .ThenBy(s => GetLocationForSorting(s).X)
                .ToList();
            
            // 2. Get existing marks if remarkAll = false
            var existingMarkNumbers = new HashSet<int>();
            if (!remarkAll)
            {
                foreach (var sleeve in sortedSleeves)
                {
                    var existing = sleeve.LookupParameter("MEP Mark")?.AsString();
                    if (!string.IsNullOrEmpty(existing))
                    {
                        var num = ExtractNumberFromMark(existing, new[] { DISCIPLINE_PREFIX, "MEP_", projectPrefix + DISCIPLINE_PREFIX });
                        if (num.HasValue) existingMarkNumbers.Add(num.Value);
                    }
                }
            }
            
            // 3. Find starting number (max of startNumber and existing max + 1)
            int nextNumber = startNumber;
            if (existingMarkNumbers.Count > 0)
            {
                nextNumber = Math.Max(nextNumber, existingMarkNumbers.Max() + 1);
            }
            
            // 4. Assign numbers
            foreach (var sleeve in sortedSleeves)
            {
                if (!remarkAll)
                {
                    var existing = sleeve.LookupParameter("MEP Mark")?.AsString();
                    if (!string.IsNullOrEmpty(existing))
                    {
                        // Keep existing - don't add to results
                        continue;
                    }
                }
                
                // Find next unused number
                while (existingMarkNumbers.Contains(nextNumber)) nextNumber++;
                existingMarkNumbers.Add(nextNumber);
                
                // Build mark string: [ProjectPrefix][DisciplinePrefix][Number]
                // Example: "SLEEVE_MEP001" or "MEP001"
                string mark = $"{projectPrefix}{DISCIPLINE_PREFIX}{nextNumber.ToString(numberFormat)}";
                
                results.Add((sleeve.Id, mark));
                nextNumber++;
            }
            
            if (!DeploymentConfiguration.DeploymentMode)
            {
                string logPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                File.AppendAllText(logPath, 
                    $"[COMBINED-MARKS] Calculated {results.Count} marks (format: {projectPrefix}{DISCIPLINE_PREFIX}{{number:{numberFormat}}})\n");
            }
            
            return results;
        }

        /// <summary>
        /// Applies calculated marks to combined sleeves in a single transaction.
        /// Optimized for BIM 360 by minimizing sync overhead.
        /// </summary>
        public (int Success, int Failed) ApplyCombinedSleeveMarksBatch(
            Document doc,
            List<(ElementId SleeveId, string Mark)> markAssignments)
        {
            if (markAssignments == null || markAssignments.Count == 0) 
                return (0, 0);
            
            int successCount = 0;
            int failedCount = 0;
            
            // Use existing transaction if available (for "Mark All" flow)
            bool localTransaction = !doc.IsModifiable;
            Transaction t = null;
            
            try
            {
                if (localTransaction)
                {
                    t = new Transaction(doc, "Apply MEP Marks (Combined Sleeves)");
                    t.Start();
                    
                    var opts = t.GetFailureHandlingOptions();
                    opts.SetFailuresPreprocessor(new Models.ParameterTransferWarningSwallower());
                    t.SetFailureHandlingOptions(opts);
                }
                
                foreach (var (sleeveId, mark) in markAssignments)
                {
                    try
                    {
                        var sleeve = doc.GetElement(sleeveId) as FamilyInstance;
                        if (sleeve == null)
                        {
                            failedCount++;
                            continue;
                        }
                        
                        var markParam = sleeve.LookupParameter("MEP Mark") 
                                     ?? sleeve.LookupParameter("MEP_Mark")
                                     ?? sleeve.LookupParameter("Mark");
                        
                        if (markParam != null && !markParam.IsReadOnly)
                        {
                            markParam.Set(mark);
                            successCount++;
                        }
                        else
                        {
                            failedCount++;
                        }
                    }
                    catch
                    {
                        failedCount++;
                    }
                }
                
                if (localTransaction && t != null)
                {
                    t.Commit();
                }
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    string logPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                    File.AppendAllText(logPath, 
                        $"[COMBINED-MARKS] ✅ Applied {successCount}/{markAssignments.Count} marks (failed: {failedCount})\n");
                }
            }
            catch (Exception ex)
            {
                if (localTransaction && t != null && t.HasStarted())
                {
                    t.RollBack();
                }
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    string logPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                    File.AppendAllText(logPath, $"[COMBINED-MARKS] ❌ Error: {ex.Message}\n");
                }
                throw;
            }
            
            return (successCount, failedCount);
        }

        private string GetLevelNameForSorting(FamilyInstance fi, Document doc)
        {
            try
            {
                var levelId = fi.LevelId;
                if (levelId != null && levelId != ElementId.InvalidElementId)
                {
                    return doc.GetElement(levelId)?.Name ?? "ZZZ";
                }
            }
            catch { }
            return "ZZZ";
        }

        private XYZ GetLocationForSorting(FamilyInstance fi)
        {
            try
            {
                if (fi.Location is LocationPoint lp)
                    return lp.Point;
            }
            catch { }
            return XYZ.Zero;
        }

        #endregion
        #region Reset Functionality

        /// <summary>
        /// ✅ RESET MARKS: Clears "MEP Mark" from sleeves on a given level.
        /// Supports "Active View Only" filtering using coordinate transformation and fallback.
        /// </summary>
        public int ResetMarksForLevel(Document doc, string levelName, BoundingBoxXYZ viewExtent = null)
        {
            var repo = new ClashZoneRepository(new SleeveDbContext());
            var allSleeves = repo.GetSleevesForLevel(levelName, null); // Get all categories

            // Filter by View Extent if provided (Session Sensitive)
            if (viewExtent != null)
            {
                // Transform view crop box to world coordinates
                Transform viewTransform = viewExtent.Transform;
                XYZ bMin = viewExtent.Min;
                XYZ bMax = viewExtent.Max;
                
                var corners = new List<XYZ>
                {
                    viewTransform.OfPoint(new XYZ(bMin.X, bMin.Y, bMin.Z)),
                    viewTransform.OfPoint(new XYZ(bMax.X, bMin.Y, bMin.Z)),
                    viewTransform.OfPoint(new XYZ(bMax.X, bMax.Y, bMin.Z)),
                    viewTransform.OfPoint(new XYZ(bMin.X, bMax.Y, bMin.Z))
                };
                
                double worldMinX = corners.Min(c => c.X) - 0.01;
                double worldMinY = corners.Min(c => c.Y) - 0.01;
                double worldMaxX = corners.Max(c => c.X) + 0.01;
                double worldMaxY = corners.Max(c => c.Y) + 0.01;

                allSleeves = allSleeves.Where(z => {
                    // ✅ COORDINATE FALLBACK (Matches ApplyMarks logic)
                    double pX = z.SleevePlacementPointActiveDocumentX;
                    double pY = z.SleevePlacementPointActiveDocumentY;
                    
                    if (Math.Abs(pX) < 0.001 && Math.Abs(pY) < 0.001)
                    {
                        pX = z.SleevePlacementPointX; // Fallback 1
                        pY = z.SleevePlacementPointY;
                    }
                    if (Math.Abs(pX) < 0.001 && Math.Abs(pY) < 0.001)
                    {
                        pX = z.IntersectionPoint.X;   // Fallback 2
                        pY = z.IntersectionPoint.Y;
                    }
                    
                    return pX >= worldMinX && pX <= worldMaxX &&
                           pY >= worldMinY && pY <= worldMaxY;
                }).ToList();
            }

            // Execute Reset in Transaction
            int resetCount = 0;
            using (var t = new Transaction(doc, "Reset MEP Marks"))
            {
                t.Start();
                foreach (var zone in allSleeves)
                {
                    ElementId id = GetElementIdSafe(zone.SleeveInstanceId);
                    if (id == ElementId.InvalidElementId) continue;

                    var fi = doc.GetElement(id) as FamilyInstance;
                    if (fi == null) continue;

                    var param = fi.LookupParameter("MEP Mark");
                    if (param != null && !param.IsReadOnly)
                    {
                        param.Set(""); // Clear mark
                        resetCount++;
                    }
                }
                t.Commit();
            }

            return resetCount;
        }

        /// <summary>
        /// ✅ RESET SELECTION: Clears "MEP Mark" from currently selected sleeves.
        /// </summary>
        public int ResetMarksForSelection(Document doc, ICollection<ElementId> selectedIds)
        {
            if (selectedIds == null || selectedIds.Count == 0) return 0;

            int resetCount = 0;
            using (var t = new Transaction(doc, "Reset Selected Marks"))
            {
                t.Start();
                foreach (var id in selectedIds)
                {
                    var ele = doc.GetElement(id);
                    if (ele == null) continue;

                    var param = ele.LookupParameter("MEP Mark");
                    if (param != null && !param.IsReadOnly)
                    {
                        param.Set("");
                        resetCount++;
                    }
                }
                t.Commit();
            }
            return resetCount;
        }

        /// <summary>
        /// ✅ RESET COUNTERS: Resets global numbering counters in database.
        /// </summary>
        public void ResetCategoryCounters(string category = null)
        {
            var markerRepo = new Data.Repositories.CategoryProcessingMarkerRepository(new SleeveDbContext());
            
            if (!string.IsNullOrEmpty(category))
            {
                markerRepo.ResetMarker(category);
            }
            else
            {
                // Reset all standard categories
                markerRepo.ResetMarker("Ducts");
                markerRepo.ResetMarker("Pipes");
                markerRepo.ResetMarker("Cable Trays");
                markerRepo.ResetMarker("Duct Accessories");
            }
        }
        #endregion
    }
}

