using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    public partial class MarkParameterService
    {
        // Implement the Batch version of ApplyMepMarkToClusters
        public (int processedCount, int errorCount) ApplyMepMarkToClustersBatch(
            Document doc, 
            string category, 
            string projectPrefix, 
            string disciplinePrefix, 
            bool remarkAll = false, 
            string numberFormat = "000", 
            MarkPrefixSettings? markPrefixes = null)
        {
            int processedCount = 0;
            int errorCount = 0;
            
            // 0. SETUP & LOGGING
            string mepmarkLogPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
            var stopwatchTotal = System.Diagnostics.Stopwatch.StartNew();
            var stopwatchRead = new System.Diagnostics.Stopwatch();
            var stopwatchEnrich = new System.Diagnostics.Stopwatch();
            var stopwatchCalc = new System.Diagnostics.Stopwatch();
            var stopwatchWrite = new System.Diagnostics.Stopwatch();

            if (!DeploymentConfiguration.DeploymentMode)
            {
               File.AppendAllText(mepmarkLogPath, $"\n===== BATCH MEPMARK SESSION STARTED {DateTime.Now} =====\n");
               File.AppendAllText(mepmarkLogPath, $"Category: {category}, ProjectPrefix: {projectPrefix}, RemarkAll: {remarkAll}\n");
            }

            EnsureSharedParametersLoaded(doc);

            // 1. READ PHASE (Main Thread)
            stopwatchRead.Start();
            // Collect all potential sleeves efficiently
            List<FamilyInstance> allSleeves;
            using (var initContext = new SleeveDbContext(doc))
            {
                allSleeves = GetAllSleevesForCategory(doc, category, initContext, markPrefixes); 
            }
            
            if (allSleeves.Count == 0)
            {
                stopwatchRead.Stop();
                if (!DeploymentConfiguration.DeploymentMode) 
                    File.AppendAllText(mepmarkLogPath, $"No sleeves found. Read Time: {stopwatchRead.ElapsedMilliseconds}ms\n");
                return (0, 0);
            }

            // Convert to lightweight identities
            var identities = new List<MarkSleeveIdentity>(allSleeves.Count);
            foreach (var sleeve in allSleeves)
            {
                var id = new MarkSleeveIdentity
                {
                    SleeveId = sleeve.Id,
                    SleeveInstanceId = sleeve.Id.IntegerValue,
                    FamilyName = sleeve.Symbol?.Family?.Name ?? "",
                    ExistingMark = sleeve.LookupParameter("MEP Mark")?.AsString() ?? sleeve.LookupParameter("Mark")?.AsString(),
                    LocationPoint = (sleeve.Location as LocationPoint)?.Point ?? XYZ.Zero,
                    LevelName = doc.GetElement(sleeve.LevelId)?.Name ?? "Unknown"
                };
                
                // Read MEP Element ID
                var mepIdParam = sleeve.LookupParameter("MEP_ElementId");
                if (mepIdParam != null) id.MepElementId = mepIdParam.AsInteger();

                // Read MEP Category (for Combined Sleeves)
                var catParam = sleeve.LookupParameter("MEP_Category");
                if (catParam != null) id.MepCategory = catParam.AsString();
                
                // ✅ Detect Combined Sleeves: If MEP_Category contains "Multi" or "Combined", OR Combined Sleeve Instance ID > 0
                var combinedIdParam = sleeve.LookupParameter("Combined Sleeve Instance ID");
                if (combinedIdParam != null && combinedIdParam.AsInteger() > 0)
                {
                    id.IsCombinedSleeve = true;
                }
                else if (!string.IsNullOrEmpty(id.MepCategory) && 
                         (id.MepCategory.Contains("Multi", StringComparison.OrdinalIgnoreCase) || 
                          id.MepCategory.Contains("Combined", StringComparison.OrdinalIgnoreCase)))
                {
                    id.IsCombinedSleeve = true;
                }

                identities.Add(id);
            }
            stopwatchRead.Stop();

            // 2. DATA ENRICHMENT (DB Calls - Optimized Bulk Load)
            stopwatchEnrich.Start();
            using (var dbContext = new SleeveDbContext(doc, msg => { }))
            {
                var repo = new ClashZoneRepository(dbContext);
                
                // Get all sleeve IDs first
                var sleeveIds = identities
                    .Where(id => !id.IsCombinedSleeve)
                    .Select(id => id.SleeveInstanceId)
                    .Distinct()
                    .ToList();
                
                // BULK LOOKUP: Get all ClashZones in one query
                // Returns map of SleeveInstanceId -> ClashZone
                // Also need to handle MEP Element ID fallback if needed, but primary key is usually SleeveInstanceId
                var zones = repo.GetClashZonesBySleeveIds(sleeveIds);
                
                // Create lookup dictionary: Handle multiple matches (should be unique per sleeve)
                // If a sleeve has multiple zones (rare/error), take the first one
                var zoneMap = zones
                    .Where(z => z.SleeveInstanceId > 0)
                    .GroupBy(z => z.SleeveInstanceId)
                    .ToDictionary(g => g.Key, g => g.First());

                // Parallel Enrichment (Filter & Prefix Data)
                var filteredBag = new System.Collections.Concurrent.ConcurrentBag<MarkSleeveIdentity>();
                
                // Use Parallel.ForEach for in-memory matching and filtering
                System.Threading.Tasks.Parallel.ForEach(identities, id => 
                {
                    // ✅ Combined Sleeves: Always include
                    if (id.IsCombinedSleeve)
                    {
                        filteredBag.Add(id);
                        return;
                    }
                    
                    ClashZone zone = null;
                    if (zoneMap.TryGetValue(id.SleeveInstanceId, out var foundZone))
                    {
                        zone = foundZone;
                    }
                    // Fallback to single ID lookup only if missing from bulk (should be rare/never)
                    // We skip the fallback here for performance - if it wasn't returned by GetClashZonesBySleeveIds, 
                    // it likely doesn't exist or isn't a clash zone sleeve.
                    
                    if (zone != null && zone.MepElementCategory == category)
                    {
                         // Match!
                         // Resolve system type from Zone params if available
                         id.SystemType = GetClashParameterValue(zone, "System Type", "MEP System Type", "System Classification");
                         id.ServiceType = GetClashParameterValue(zone, "Service Type", "System Abbreviation", "MEP System Type");
                         filteredBag.Add(id);
                    }
                });
                
                identities = filteredBag.ToList();
            }
            stopwatchEnrich.Stop();
            if (identities.Count == 0) return (0, 0);

            // 3. CALCULATE PHASE (Moved to PrepareMepMarkAssignments)
            stopwatchCalc.Start();
            var sortedIdentities = PrepareMepMarkAssignments(doc, category, projectPrefix, disciplinePrefix, remarkAll, numberFormat, markPrefixes);
            stopwatchCalc.Stop();

            // 4. WRITE PHASE
            stopwatchWrite.Start();
            var (processed, errors) = ApplyMarkAssignments(doc, category, sortedIdentities);
            stopwatchWrite.Stop();

            // Update Marker (DB)
            UpdateCategoryMarkerFromAssignments(doc, category, sortedIdentities, disciplinePrefix, markPrefixes);
            
            stopwatchTotal.Stop();
            LogBatchPerformance(category, stopwatchTotal, stopwatchRead, stopwatchEnrich, stopwatchCalc, stopwatchWrite, processed);

            return (processed, errors);
        }

        /// <summary>
        /// ✅ NEW: Prepares mark assignments for a category without writing to Revit.
        /// This is safe to run in parallel across categories (Read phase must be on main thread).
        /// </summary>
        public List<MarkSleeveIdentity> PrepareMepMarkAssignments(
            Document doc, 
            string category, 
            string projectPrefix, 
            string disciplinePrefix, 
            bool remarkAll = false, 
            string numberFormat = "000", 
            MarkPrefixSettings? markPrefixes = null)
        {
            // 1. READ PHASE (Main Thread required)
            List<FamilyInstance> allSleeves;
            using (var initContext = new SleeveDbContext(doc))
            {
                allSleeves = GetAllSleevesForCategory(doc, category, initContext, markPrefixes); 
            }
            
            return PrepareMepMarkAssignments_WithSleeves(doc, category, allSleeves, projectPrefix, disciplinePrefix, remarkAll, numberFormat, markPrefixes);
        }

        /// <summary>
        /// ✅ NEW: Internal preparation logic that can run on any thread once sleeves are collected.
        /// </summary>
        public List<MarkSleeveIdentity> PrepareMepMarkAssignments_WithSleeves(
            Document doc,
            string category,
            List<FamilyInstance> allSleeves,
            string projectPrefix, 
            string disciplinePrefix, 
            bool remarkAll = false, 
            string numberFormat = "000", 
            MarkPrefixSettings? markPrefixes = null)
        {
            if (allSleeves == null || allSleeves.Count == 0) return new List<MarkSleeveIdentity>();

            var identities = new List<MarkSleeveIdentity>(allSleeves.Count);
            foreach (var sleeve in allSleeves)
            {
                // Note: Accessing basic properties like Id, Symbol.Family.Name, LevelId is usually safe 
                // but LookupParameter might be slightly riskier. However, for read-only it's generally fine.
                var id = new MarkSleeveIdentity
                {
                    SleeveId = sleeve.Id,
                    SleeveInstanceId = sleeve.Id.IntegerValue,
                    FamilyName = sleeve.Symbol?.Family?.Name ?? "",
                    ExistingMark = sleeve.LookupParameter("MEP Mark")?.AsString() ?? sleeve.LookupParameter("Mark")?.AsString(),
                    LocationPoint = (sleeve.Location as LocationPoint)?.Point ?? XYZ.Zero,
                    LevelName = doc.GetElement(sleeve.LevelId)?.Name ?? "Unknown"
                };
                
                var mepIdParam = sleeve.LookupParameter("MEP_ElementId");
                if (mepIdParam != null) id.MepElementId = mepIdParam.AsInteger();

                var catParam = sleeve.LookupParameter("MEP_Category");
                if (catParam != null) id.MepCategory = catParam.AsString();
                
                var combinedIdParam = sleeve.LookupParameter("Combined Sleeve Instance ID");
                if (combinedIdParam != null && combinedIdParam.AsInteger() > 0) id.IsCombinedSleeve = true;
                else if (!string.IsNullOrEmpty(id.MepCategory) && 
                         (id.MepCategory.Contains("Multi", StringComparison.OrdinalIgnoreCase) || 
                          id.MepCategory.Contains("Combined", StringComparison.OrdinalIgnoreCase)))
                {
                    id.IsCombinedSleeve = true;
                }

                identities.Add(id);
            }

            // 2. DATA ENRICHMENT (Parallel DB Calls)
            using (var dbContext = new SleeveDbContext(doc, msg => { }))
            {
                var repo = new ClashZoneRepository(dbContext);
                var sleeveIds = identities.Where(id => !id.IsCombinedSleeve).Select(id => id.SleeveInstanceId).Distinct().ToList();
                var zoneMap = repo.GetClashZonesBySleeveIds(sleeveIds)
                    .Where(z => z.SleeveInstanceId > 0)
                    .GroupBy(z => z.SleeveInstanceId)
                    .ToDictionary(g => g.Key, g => g.First());

                var filteredBag = new System.Collections.Concurrent.ConcurrentBag<MarkSleeveIdentity>();
                System.Threading.Tasks.Parallel.ForEach(identities, id => 
                {
                    if (id.IsCombinedSleeve) { filteredBag.Add(id); return; }
                    
                    if (zoneMap.TryGetValue(id.SleeveInstanceId, out var zone) && zone.MepElementCategory == category)
                    {
                         id.SystemType = GetClashParameterValue(zone, "System Type", "MEP System Type", "System Classification");
                         id.ServiceType = GetClashParameterValue(zone, "Service Type", "System Abbreviation", "MEP System Type");
                         filteredBag.Add(id);
                    }
                });
                identities = filteredBag.ToList();
            }

            if (identities.Count == 0) return identities;

            // 3. CALCULATE PHASE
            var sortedIdentities = identities.OrderBy(x => x.LevelName)
                                             .ThenByDescending(x => x.LocationPoint.Y)
                                             .ThenBy(x => x.LocationPoint.X)
                                             .ToList();

            var candidatePrefixes = GetCandidateDisciplinePrefixes(category, disciplinePrefix, markPrefixes);
            int categoryMaxNumber = GetMaxExistingMarkNumberForCategory(doc, category, candidatePrefixes);
            int startCount = GetCategoryMarkerCount(doc, category);
            int nextNumber = Math.Max(categoryMaxNumber, startCount) + 1;
            
            var usedNumbers = remarkAll ? GetExistingMarkNumbers(doc, category, candidatePrefixes) : new HashSet<int>();

            // ✅ TASK DIVISION: Parallel prefix resolution
            System.Threading.Tasks.Parallel.ForEach(sortedIdentities, id => {
                if (id.IsCombinedSleeve) id.CalculatedPrefix = "MEP";
                else {
                    var pfx = markPrefixes != null ? ResolveDisciplinePrefixFromStrings(category, disciplinePrefix, markPrefixes, id.SystemType, id.ServiceType) : disciplinePrefix;
                    id.CalculatedPrefix = string.IsNullOrEmpty(pfx) ? disciplinePrefix : pfx;
                }
            });

            // Sequential Numbering (Must be sequential to avoid collisions/gaps)
            foreach (var id in sortedIdentities)
            {
                if (!remarkAll && !string.IsNullOrEmpty(id.ExistingMark)) continue;

                int numberToUse = nextNumber;
                if (remarkAll && !string.IsNullOrEmpty(id.ExistingMark))
                {
                    int? extracted = ExtractNumberFromMark(id.ExistingMark, candidatePrefixes);
                    if (extracted.HasValue) { numberToUse = extracted.Value; usedNumbers.Add(numberToUse); }
                    else { while(usedNumbers.Contains(nextNumber)) nextNumber++; numberToUse = nextNumber++; usedNumbers.Add(numberToUse); }
                }
                else { while(usedNumbers.Contains(nextNumber)) nextNumber++; numberToUse = nextNumber++; usedNumbers.Add(numberToUse); }
                
                id.CalculatedMark = $"{projectPrefix}{id.CalculatedPrefix}{numberToUse.ToString(numberFormat)}";
            }

            return sortedIdentities;
        }

        private int GetCategoryMarkerCount(Document doc, string category)
        {
            try {
                using var markerContext = new SleeveDbContext(doc);
                var markerRepo = new CategoryProcessingMarkerRepository(markerContext, null);
                return markerRepo.GetMarker(category).lastCount;
            } catch { return 0; }
        }

        public void UpdateCategoryMarkerFromAssignments(Document doc, string category, List<MarkSleeveIdentity> assignments, string disciplinePrefix, MarkPrefixSettings? markPrefixes)
        {
            try {
                int maxAssignedNumber = 0;
                var candidatePrefixes = GetCandidateDisciplinePrefixes(category, disciplinePrefix, markPrefixes);
                foreach (var id in assignments)
                {
                    if (!string.IsNullOrEmpty(id.CalculatedMark))
                    {
                        int? num = ExtractNumberFromMark(id.CalculatedMark, candidatePrefixes);
                        if (num.HasValue && num.Value > maxAssignedNumber) maxAssignedNumber = num.Value;
                    }
                }
                if (maxAssignedNumber <= 0) return;
                using var markerContext = new SleeveDbContext(doc);
                var markerRepo = new CategoryProcessingMarkerRepository(markerContext, null);
                var info = markerRepo.GetMarker(category);
                if (maxAssignedNumber > info.lastCount) markerRepo.UpdateMarker(category, maxAssignedNumber);
            } catch {}
        }

        public (int processed, int errors) ApplyMarkAssignments(Document doc, string category, List<MarkSleeveIdentity> assignments)
        {
            int processed = 0, errors = 0;
            bool localTransaction = !doc.IsModifiable;
            Transaction t = null;
            try {
                if (localTransaction) { 
                    t = new Transaction(doc, "Apply MEP Marks (" + category + ")"); 
                    t.Start(); 
                    var opts = t.GetFailureHandlingOptions();
                    opts.SetFailuresPreprocessor(new JSE_RevitAddin_MEP_OPENINGS.Models.ParameterTransferWarningSwallower());
                    t.SetFailureHandlingOptions(opts);
                }
                foreach (var id in assignments) {
                    if (string.IsNullOrEmpty(id.CalculatedMark)) continue;
                    try {
                        var el = doc.GetElement(id.SleeveId);
                        if (el != null) { SetMarkParameter(el, id.CalculatedMark); processed++; }
                    } catch { errors++; }
                }
                if (localTransaction) t.Commit();
            } finally { if (t != null) t.Dispose(); }
            return (processed, errors);
        }

        private void LogBatchPerformance(string cat, System.Diagnostics.Stopwatch total, System.Diagnostics.Stopwatch r, System.Diagnostics.Stopwatch e, System.Diagnostics.Stopwatch c, System.Diagnostics.Stopwatch w, int count)
        {
            if (DeploymentConfiguration.DeploymentMode) return;
            string log = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
            var msg = $"PERFORMANCE [{cat}]: Total={total.ElapsedMilliseconds}ms | R={r.ElapsedMilliseconds}ms | E={e.ElapsedMilliseconds}ms | C={c.ElapsedMilliseconds}ms | W={w.ElapsedMilliseconds}ms ({count} items)\n";
            File.AppendAllText(log, msg);
        }
    }
}
