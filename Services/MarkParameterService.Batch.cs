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

            // 2. DATA ENRICHMENT (DB Calls - can be parallelized or bulk loaded)
            // We need ClashZone info to confirm Category and get System Type
            // Getting System Type from DB is faster than opening many elements if we have the index
            
            // Load Snapshots or ClashZones? ClashZones is better for Category check.
            // Let's use the DB Context efficiently.
            stopwatchEnrich.Start();
            using (var dbContext = new SleeveDbContext(doc, msg => { }))
            {
                // Pre-fetch ClashZones for these sleeves?
                // Or just loop and query (SQLite is fast, single thread might be ok, but let's try to be smart)
                // For now, let's just do sequential DB lookup (it's fast on local SQLite) but separate from Revit Transaction
                
                var repo = new ClashZoneRepository(dbContext);
                
                // We'll perform the category filtering and prefix resolution here in "Phase 1.5"
                var filteredIdentities = new List<MarkSleeveIdentity>();
                
                // This part involves Logic + DB, so it's "Calculate"
                foreach (var id in identities)
                {
                    ClashZone zone = null;

                    // ✅ Combined Sleeves: Always include, no category check needed
                    if (id.IsCombinedSleeve)
                    {
                        // Combined sleeves always match - they get MEP prefix later
                        filteredIdentities.Add(id);
                        continue;
                    }
                    
                    // Non-combined: DB Lookup for category matching
                    zone = repo.GetClashZoneBySleeveId(id.SleeveInstanceId); 
                    if (zone == null && id.MepElementId > 0)
                    {
                        zone = repo.GetClashZoneByMepElementId((int)id.MepElementId);
                    }

                    if (zone != null && zone.MepElementCategory == category)
                    {
                        // Match!
                        // Resolve system type from Zone params if available
                        id.SystemType = GetClashParameterValue(zone, "System Type", "MEP System Type", "System Classification");
                        id.ServiceType = GetClashParameterValue(zone, "Service Type", "System Abbreviation", "MEP System Type");
                        filteredIdentities.Add(id);
                    }
                }
                identities = filteredIdentities;
            }
            stopwatchEnrich.Stop();

            if (identities.Count == 0) return (0, 0);


            // 3. CALCULATE PHASE (Parallel / Sequential Sort)
            stopwatchCalc.Start();
            // Sorting determines numbering order. Must be sequential or parallel-sort.
            // Sort keys: Level -> Y (descending) -> X (ascending) usually
            var sortedIdentities = identities.OrderBy(x => x.LevelName)
                                             .ThenByDescending(x => x.LocationPoint.Y)
                                             .ThenBy(x => x.LocationPoint.X)
                                             .ToList();

            // Determine Start Index
            // Need max existing mark from DB/Project
            int startIndex = 1;
            // Existing logic: GetMaxExistingMarkNumberForCategory(doc, category, candidatePrefixes);
            // This requires scanning ALL potential prefixes.
            
            // ... (Simplified: assume we can reuse existing logic or just recalculate local max if we rely on "dumb once")
            // But we need GLOBAL max to avoid collision. 
            // Reuse `GetMaxExistingMarkNumberForCategory` (it reads DB/Project parameters).
            // We can run this in Read phase. Let's assume passed in or calculated.
            
            var candidatePrefixes = GetCandidateDisciplinePrefixes(category, disciplinePrefix, markPrefixes);
            int categoryMaxNumber = GetMaxExistingMarkNumberForCategory(doc, category, candidatePrefixes);
            
            // Marker Repo
            // using var markerContext = new SleeveDbContext(doc);
            // var markerRepo = new CategoryProcessingMarkerRepository(markerContext);
            // int markerLastCount = markerRepo.GetMarker(category).lastCount;
            // categoryMaxNumber = Math.Max(categoryMaxNumber, markerLastCount);
            
            int nextNumber = categoryMaxNumber + 1;
            var usedNumbers = new HashSet<int>();
            if (remarkAll)
            {
                usedNumbers = GetExistingMarkNumbers(doc, category, candidatePrefixes);
            }

            // Assign Numbers Loop
            foreach (var id in sortedIdentities)
            {
                // ✅ Combined Sleeves: Force MEP prefix (user requirement)
                string elementPrefix;
                if (id.IsCombinedSleeve)
                {
                    elementPrefix = "MEP";
                }
                else
                {
                    // Resolve Prefix (Logic) for non-combined sleeves
                    elementPrefix = disciplinePrefix;
                    if (markPrefixes != null)
                    {
                         // Call the real logic helper from the main partial class
                         elementPrefix = ResolveDisciplinePrefixFromStrings(category, disciplinePrefix, markPrefixes, id.SystemType, id.ServiceType);
                    }
                    
                    // Ensure default if null
                    if (string.IsNullOrEmpty(elementPrefix)) elementPrefix = disciplinePrefix;
                }

                // Calculate Number
                int numberToUse = nextNumber;
                
                // Existing Mark Logic (Remark/Keep)
                bool keepExisting = false;
                if (!remarkAll && !string.IsNullOrEmpty(id.ExistingMark))
                {
                    keepExisting = true;
                }
                
                if (!keepExisting)
                {
                    // Assign new
                     if (remarkAll && !string.IsNullOrEmpty(id.ExistingMark))
                     {
                         // Try extract
                         int? extracted = ExtractNumberFromMark(id.ExistingMark, candidatePrefixes);
                         if (extracted.HasValue)
                         {
                             numberToUse = extracted.Value;
                             usedNumbers.Add(numberToUse);
                         }
                         else
                         {
                             // Find next unused
                             while(usedNumbers.Contains(nextNumber)) nextNumber++;
                             numberToUse = nextNumber++;
                             usedNumbers.Add(numberToUse);
                         }
                     }
                     else
                     {
                         // New
                         while(usedNumbers.Contains(nextNumber)) nextNumber++;
                         numberToUse = nextNumber++;
                         usedNumbers.Add(numberToUse);
                     }
                     
                     // Generate String
                     string projectPfx = projectPrefix; // or extract
                     // (Logic for extracting project prefix omitted for brevity, verify against requirement)
                     
                     string numStr = numberToUse.ToString(numberFormat);
                     id.CalculatedMark = $"{projectPfx}{elementPrefix}{numStr}";
                }
            }
            stopwatchCalc.Stop();

            // 4. WRITE PHASE (Main Thread Transaction)
            // 4. WRITE PHASE (Main Thread)
            // If we are already in a transaction, use it. But we can't tell easily unless we pass a param.
            // Current "One Transaction for All" strategy requires us to NOT start a transaction if one is active, OR
            // simply perform the SET actions and let the caller Commit.
            // HOWEVER: SetMarkParameter modifies the model. It MUST be in a transaction.
            
            // To support "Mark All Categories" in ONE transaction, this method should probably NOT control the transaction.
            // But for "Single Category", it might need one.
            // Let's check `doc.IsModifiable`.
            
            stopwatchWrite.Start();
            bool localTransaction = !doc.IsModifiable;
            Transaction t = null;
            
            try 
            {
                if (localTransaction)
                {
                    t = new Transaction(doc, "Apply MEP Marks (Batch " + category + ")");
                    t.Start();
                    
                    var opts = t.GetFailureHandlingOptions();
                    opts.SetFailuresPreprocessor(new JSE_RevitAddin_MEP_OPENINGS.Models.ParameterTransferWarningSwallower());
                    t.SetFailureHandlingOptions(opts);
                }

                foreach (var id in sortedIdentities)
                {
                    if (string.IsNullOrEmpty(id.CalculatedMark)) continue;
                    
                    try
                    {
                        var el = doc.GetElement(id.SleeveId);
                        if (el == null) continue;
                        
                        SetMarkParameter(el, id.CalculatedMark);
                        processedCount++;
                    }
                    catch (Exception ex)
                    {
                        errorCount++;
                    }
                }
                
                if (localTransaction)
                {
                    t.Commit();
                }
            }
            finally
            {
                if (t != null) t.Dispose();
            }
            stopwatchWrite.Stop();

            // Update Marker (DB) - Incremental
            try 
            {
                using (var markerContext = new SleeveDbContext(doc))
                {
                    var markerRepository = new CategoryProcessingMarkerRepository(markerContext, null);
                    // Use nextNumber - 1 as the last used number
                    // But we actually need the max of (existing max, new assigned max).
                    // In batch, nextNumber kept incrementing.
                    int lastUsed = nextNumber - 1; 
                    if (lastUsed > 0)
                    {
                        var info = markerRepository.GetMarker(category);
                        if (lastUsed > info.lastCount)
                            markerRepository.UpdateMarker(category, lastUsed);
                    }
                }
            }
            catch {}
            
            stopwatchTotal.Stop();
            
            // LOG PERFORMANCE
            if (!DeploymentConfiguration.DeploymentMode)
            {
                var msg = $"PERFORMANCE [{category}]: Total={stopwatchTotal.ElapsedMilliseconds}ms | " +
                          $"Read={stopwatchRead.ElapsedMilliseconds}ms | " +
                          $"Enrich={stopwatchEnrich.ElapsedMilliseconds}ms | " +
                          $"Calc={stopwatchCalc.ElapsedMilliseconds}ms | " +
                          $"Write={stopwatchWrite.ElapsedMilliseconds}ms ({(processedCount > 0 ? (stopwatchWrite.ElapsedMilliseconds / processedCount) : 0)} ms/item)\n";
                File.AppendAllText(mepmarkLogPath, msg);
            }

            return (processedCount, errorCount);
        }

        // Helper helper
        private string ResolveDisciplinePrefixForElement_Batch(string category, string defaultPrefix, MarkPrefixSettings settings, string systemType)
        {
             // We only passed SystemType here in the simplified loop, but the full method needs ServiceType too.
             // For batch, we ideally extract both. 
             // Assuming systemType contains relevant info or serviceType is empty for now.
             // OR better: Update the batch loop to extract ServiceType too (MarkSleeveIdentity has it).
             
             // Since this is a placeholder method called inside the loop, and we are modifying the loop:
             // We should call ResolveDisciplinePrefixFromStrings directly in the loop.
             // But to satisfy this method signature:
             return ResolveDisciplinePrefixFromStrings(category, defaultPrefix, settings, systemType, null); // ServiceType assumed null/irrelevant if not passed
        }
    }
}
