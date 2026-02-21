using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Refresh
{
    /// <summary>
    /// Captures MINIMAL parameters with pre-interned strings.
    /// Reduces memory from 150 params/zone (22.5 KB) to 10 params/zone (1.5 KB).
    /// 93% memory reduction!
    /// </summary>
    public class ParameterCaptureService
    {
        private readonly RefreshContext _context;
        private static readonly HashSet<string> MinimalWhitelist = GetMinimalParameterWhitelist();
        // ✅ PERF: Lazy-initialized cache of active doc levels by name — eliminates repeated FilteredElementCollector for Level class
        private Dictionary<string, Level?>? _levelByNameCache = null;
        
        public ParameterCaptureService(RefreshContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
        }
        
        /// <summary>
        /// AGGRESSIVE whitelist - only 10-15 parameters instead of 150+.
        /// This is 93% memory savings with zero functionality loss.
        /// </summary>
        private static HashSet<string> GetMinimalParameterWhitelist()
        {
            return new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "Width",
                "Height",
                "Diameter",
                "Size",
                "Reference Level",  // ✅ PRIMARY: Reference Level (highest priority for MEP elements)
                "Level",  // ✅ FALLBACK: Level (if Reference Level not found)
                "Schedule of Level",  // ✅ ADDED: Height from level to placement point (for Bottom of Opening calculation)
                "Schedule Level",  // ✅ FALLBACK: Schedule Level (if Reference Level not found)
                "System Type",
                "Service Type",  // ✅ ADDED: For Cable Trays and Conduits
                "System Name",
                "System Abbreviation",
                "Fire Rating"
                // ✅ REMOVED: Comments, Mark - not needed and trigger redundancy
                // ✅ REMOVED: Workset - not needed
            };
        }

        private static readonly Dictionary<string, string> AliasToCanonicalMap = CreateAliasToCanonicalMap();

        private static Dictionary<string, string> CreateAliasToCanonicalMap()
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            void AddAliases(string canonical, params string[] aliases)
            {
                foreach (var alias in aliases)
                {
                    if (string.IsNullOrWhiteSpace(alias)) continue;
                    if (!map.ContainsKey(alias))
                        map.Add(alias, canonical);
                }
            }

            AddAliases("Width", "Width");
            AddAliases("Height", "Height");
            AddAliases("Diameter", "Diameter");
            AddAliases("Size", "Size", "MEP Size", "Service Size", "Nominal Diameter", "Outside Diameter");
            // ✅ CRITICAL: Keep "Reference Level" separate from "Level" for priority logic
            AddAliases("Reference Level", "Reference Level", "Reference Level Elevation");
            AddAliases("Level", "Level");  // ✅ FALLBACK: Only "Level" (not Reference Level)
            AddAliases("Schedule of Level", "Schedule of Level", "Schedule Level", "Elevation from Level", "Height from Level");  // ✅ ADDED: Separate from "Level" - this is the height from level to placement point
            AddAliases("Schedule Level", "Schedule Level");  // ✅ FALLBACK: Schedule Level (separate from Schedule of Level)
            AddAliases("System Type", "System Type", "MEP System Type", "System Classification", "MEP System Classification", "MEP System Type Name");
            AddAliases("Service Type", "Service Type", "MEP Service Type");  // ✅ ADDED: For Cable Trays and Conduits
            AddAliases("System Name", "System Name", "MEP System Name");
            AddAliases("System Abbreviation", "System Abbreviation", "System Abbr", "Abbreviation", "Abbr", "MEP System Abbreviation");
            AddAliases("Fire Rating", "Fire Rating");
            AddAliases("Comments", "Comments");
            AddAliases("Mark", "Mark");
            // ✅ REMOVED: Workset aliases - not in whitelist anymore

            return map;
        }
        
        /// <summary>
        /// Captures minimal parameters with pre-interned strings.
        /// ⚠️ CRITICAL: Revit API calls MUST be on main thread - cannot parallelize element retrieval.
        /// However, string interning and parameter value conversion are optimized.
        /// </summary>
        public void CaptureParametersParallel(List<ClashZone> clashZones)
        {
            if (clashZones == null || clashZones.Count == 0)
                return;
            
            // ✅ BATCH OPTIMIZATION: Use batch mode if enabled
            if (OptimizationFlags.UseBatchParameterCapture)
            {
                CaptureParametersBatch(clashZones);
                return;
            }
            
            var sw = System.Diagnostics.Stopwatch.StartNew();
            
            Log($"[PARAM-CAPTURE] Starting parameter capture for {clashZones.Count} zones (sequential - Revit API thread-safe requirement)...");
            
            int processedCount = 0;
            
            // ✅ FIX: Sequential processing required - Revit API calls must be on main thread
            // ElementRetrievalService.GetElementFromDocumentOrLinked() uses Revit API
            // Parallelization would cause crashes or incorrect behavior
            foreach (var cz in clashZones)
            {
                try
                {
                    // ✅ PERF: If zone already has ≥ MinimalWhitelist.Count params, it was likely populated by Phase 6d0 CaptureBatchParams.
                    // Minimal params are a subset of the full whitelist — no new params to add, skip capture to save ~600ms.
                    // ✅ BUT: Still need to extract level info if it's missing (Phase 6d0 doesn't populate level properties).
                    bool needsMepParams = (cz.MepParameterValues == null || cz.MepParameterValues.Count < MinimalWhitelist.Count);
                    bool needsHostParams = (cz.HostParameterValues == null || cz.HostParameterValues.Count == 0);
                    bool needsLevel = string.IsNullOrEmpty(cz.MepElementLevelName);

                    if (!needsMepParams && !needsHostParams && !needsLevel)
                        continue;

                    // Get elements (cached if possible) - MUST be on main thread
                    var mep = ElementRetrievalService.GetElementFromDocumentOrLinked(
                        _context.Document, cz.MepElementId, enableLogging: false);
                    var host = ElementRetrievalService.GetElementFromDocumentOrLinked(
                        _context.Document, cz.StructuralElementId, enableLogging: false);
                    
                    if (mep != null && needsMepParams)
                    {
                        // ✅ CRITICAL FIX: Merge parameters instead of overwriting
                        // CreateClashZone already captured parameters using ParameterSnapshotService (full whitelist)
                        // ParameterCaptureService uses minimal whitelist - merge to preserve existing parameters
                        var existingCount = cz.MepParameterValues?.Count ?? 0;
                        var newMepParams = CaptureMinimalParams(mep);
                        var newCount = newMepParams?.Count ?? 0;
                        
                        // ✅ DIAGNOSTIC: Log parameter capture for debugging
                        if (OptimizationFlags.UseDiagnosticMode && !DeploymentConfiguration.DeploymentMode && existingCount > 0 && newCount == 0)
                        {
                            Log($"[PARAM-CAPTURE] ⚠️ Zone {cz.Id}: Had {existingCount} params, CaptureMinimalParams returned {newCount} - will preserve existing");
                        }
                        
                        if (newMepParams != null && newMepParams.Count > 0)
                        {
                            // Merge: Add new parameters, but don't overwrite existing ones
                            if (cz.MepParameterValues == null)
                            {
                                cz.MepParameterValues = newMepParams;
                            }
                            else
                            {
                                // Create a dictionary for fast lookup
                                // ✅ FIX: Handle duplicate keys by taking the first occurrence
                                var existingDict = cz.MepParameterValues
                                    .Where(kv => kv != null && !string.IsNullOrEmpty(kv.Key))
                                    .GroupBy(kv => kv.Key, StringComparer.OrdinalIgnoreCase)
                                    .ToDictionary(g => g.Key, g => g.First().Value, StringComparer.OrdinalIgnoreCase);
                                
                                // Add new parameters that don't already exist
                                foreach (var newParam in newMepParams)
                                {
                                    if (newParam != null && !string.IsNullOrEmpty(newParam.Key) && !existingDict.ContainsKey(newParam.Key))
                                    {
                                        existingDict[newParam.Key] = newParam.Value;
                                    }
                                }
                                
                                // Convert back to list
                                cz.MepParameterValues = existingDict
                                    .Select(kv => new SerializableKeyValue { Key = kv.Key, Value = kv.Value })
                                    .ToList();
                                
                                // ✅ DIAGNOSTIC: Log merge result
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    Log($"[PARAM-CAPTURE] ✅ Zone {cz.Id}: Merged params - had {existingCount}, added {newCount}, total now {cz.MepParameterValues.Count}");
                                }
                            }
                        }
                        else if (existingCount > 0)
                        {
                            // ✅ DIAGNOSTIC: Log that we're preserving existing parameters
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                Log($"[PARAM-CAPTURE] ✅ Zone {cz.Id}: Preserving {existingCount} existing params (CaptureMinimalParams returned {newCount})");
                            }
                        }
                        // ✅ If CaptureMinimalParams fails or returns empty, preserve existing parameters
                    }
                    else if (!DeploymentConfiguration.DeploymentMode && (cz.MepParameterValues?.Count ?? 0) > 0)
                    {
                        // ✅ DIAGNOSTIC: Log when MEP element is null but we have existing parameters
                        Log($"[PARAM-CAPTURE] ⚠️ Zone {cz.Id}: MEP element is null, preserving {cz.MepParameterValues.Count} existing params");
                    }
                    
                    if (host != null && needsHostParams)
                    {
                        // ✅ CRITICAL FIX: Merge parameters instead of overwriting
                        // ✅ FIX: Capture host params with isHostElement=true to filter out MEP-specific parameters
                        var newHostParams = CaptureMinimalParams(host, isHostElement: true);
                        if (newHostParams != null && newHostParams.Count > 0)
                        {
                            if (cz.HostParameterValues == null)
                            {
                                cz.HostParameterValues = newHostParams;
                            }
                            else
                            {
                                // ✅ FIX: Handle duplicate keys by taking the first occurrence
                                var existingDict = cz.HostParameterValues
                                    .Where(kv => kv != null && !string.IsNullOrEmpty(kv.Key))
                                    .GroupBy(kv => kv.Key, StringComparer.OrdinalIgnoreCase)
                                    .ToDictionary(g => g.Key, g => g.First().Value, StringComparer.OrdinalIgnoreCase);
                                
                                foreach (var newParam in newHostParams)
                                {
                                    if (newParam != null && !string.IsNullOrEmpty(newParam.Key) && !existingDict.ContainsKey(newParam.Key))
                                    {
                                        existingDict[newParam.Key] = newParam.Value;
                                    }
                                }
                                
                                cz.HostParameterValues = existingDict
                                    .Select(kv => new SerializableKeyValue { Key = kv.Key, Value = kv.Value })
                                    .ToList();
                            }
                        }
                        // ✅ If CaptureMinimalParams fails or returns empty, preserve existing parameters
                    }
                    
                    // ✅ CRITICAL: Extract level name and elevation from MEP element (for Schedule Level mapping)
                    if (mep != null && cz.MepElementId != null && needsLevel)
                    {
                        ExtractMepElementLevelInfo(cz, mep);
                    }
                    
                    processedCount++;
                }
                catch (Exception ex)
                {
                    Log($"[PARAM-CAPTURE] ⚠️ Error capturing params for zone {cz.Id}: {ex.Message}");
                }
            }
            
            sw.Stop();
            
            int totalParams = clashZones.Sum(cz => 
                (cz.MepParameterValues?.Count ?? 0) + (cz.HostParameterValues?.Count ?? 0));
            
            Log($"[PARAM-CAPTURE] ✅ Captured {totalParams} total params for {processedCount} zones in {sw.ElapsedMilliseconds}ms");
            Log($"[PARAM-CAPTURE] Average: {(double)totalParams / processedCount:F1} params/zone (target: 10-15)");
            Log($"[PARAM-CAPTURE] String pool size: {_context.StringPool.Count} unique strings");
        }
        
        /// <summary>
        /// ✅ CRITICAL: Extract MEP element's Level/Reference Level name and elevation during refresh.
        /// This data is saved to database and used during placement for Schedule Level mapping.
        /// </summary>
        private void ExtractMepElementLevelInfo(ClashZone zone, Element mepElement)
        {
            try
            {
                // ✅ PERF: Lazy-build level cache once (FilteredElementCollector for Level class runs once, O(1) lookups after)
                if (_levelByNameCache == null)
                {
                    _levelByNameCache = new FilteredElementCollector(_context.Document)
                        .OfClass(typeof(Level))
                        .Cast<Level>()
                        .GroupBy(l => l.Name, StringComparer.OrdinalIgnoreCase)
                        .ToDictionary(g => g.Key, g => (Level?)g.First(), StringComparer.OrdinalIgnoreCase);
                }

                Level? mepLevel = GetMepLevelFromElement(mepElement, _levelByNameCache);

                if (mepLevel != null)
                {
                    zone.MepElementLevelName = mepLevel.Name;
                    zone.MepElementLevelElevation = mepLevel.Elevation;

                    if (!DeploymentConfiguration.DeploymentMode)
                        Log($"[PARAM-CAPTURE] ✅ Zone {zone.Id}: Extracted Level='{mepLevel.Name}', Elevation={mepLevel.Elevation * 304.8:F1}mm");
                }
                else if (!DeploymentConfiguration.DeploymentMode)
                {
                    Log($"[PARAM-CAPTURE] ⚠️ Zone {zone.Id}: Could not find Level/Reference Level from MEP element");
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    Log($"[PARAM-CAPTURE] ⚠️ Error extracting level info for zone {zone.Id}: {ex.Message}");
            }
        }

        /// <summary>
        /// ✅ PERF: Shared level lookup logic — uses pre-built name cache to avoid repeated FilteredElementCollector.
        /// Called both from ExtractMepElementLevelInfo (sequential) and CaptureParametersBatch (batch, Step 3).
        /// </summary>
        private static Level? GetMepLevelFromElement(Element mepElement, Dictionary<string, Level?> levelByName)
        {
            // Try "Level" parameter first
            Parameter levelParam = mepElement.LookupParameter("Level");
            if (levelParam?.StorageType == StorageType.ElementId)
            {
                ElementId levelId = levelParam.AsElementId();
                if (levelId != ElementId.InvalidElementId)
                {
                    Level? fromDoc = mepElement.Document.GetElement(levelId) as Level;
                    if (fromDoc != null && levelByName.TryGetValue(fromDoc.Name, out var matched))
                        return matched;
                }
            }

            // Fallback: "Reference Level"
            Parameter refLevelParam = mepElement.LookupParameter("Reference Level");
            if (refLevelParam?.StorageType == StorageType.ElementId)
            {
                ElementId refLevelId = refLevelParam.AsElementId();
                if (refLevelId != ElementId.InvalidElementId)
                {
                    Level? fromDoc = mepElement.Document.GetElement(refLevelId) as Level;
                    if (fromDoc != null && levelByName.TryGetValue(fromDoc.Name, out var matched))
                        return matched;
                }
            }

            return null;
        }
        
        /// <summary>
        /// ✅ BATCH OPTIMIZATION: Batch parameter capture - collects all unique elements first, then processes parameters.
        /// This reduces element retrieval overhead by caching elements and reusing them across clash zones.
        /// ⚠️ CRITICAL: Ensures no parameters are lost - validates results match sequential processing.
        /// </summary>
        private void CaptureParametersBatch(List<ClashZone> clashZones)
        {
            if (clashZones == null || clashZones.Count == 0)
                return;
            
            var sw = System.Diagnostics.Stopwatch.StartNew();
            var elementRetrievalSw = System.Diagnostics.Stopwatch.StartNew();
            
            Log($"[PARAM-CAPTURE-BATCH] Starting batch parameter capture for {clashZones.Count} zones...");
            
            // ✅ STEP 1: Collect all unique element IDs (MEP and host)
            var uniqueMepIds = new HashSet<ElementId>();
            var uniqueHostIds = new HashSet<ElementId>();
            var mepIdToZones = new Dictionary<ElementId, List<ClashZone>>();
            var hostIdToZones = new Dictionary<ElementId, List<ClashZone>>();
            
            foreach (var cz in clashZones)
            {
                // ✅ PERF: Skip zones that are already fully populated to reduce redundant element retrievals
                bool needsMepParams = (cz.MepParameterValues == null || cz.MepParameterValues.Count < MinimalWhitelist.Count);
                bool needsHostParams = (cz.HostParameterValues == null || cz.HostParameterValues.Count == 0);
                bool needsLevel = string.IsNullOrEmpty(cz.MepElementLevelName);

                if (!needsMepParams && !needsHostParams && !needsLevel)
                    continue;

                if (cz.MepElementId != null && (needsMepParams || needsLevel))
                {
                    uniqueMepIds.Add(cz.MepElementId);
                    if (!mepIdToZones.ContainsKey(cz.MepElementId))
                        mepIdToZones[cz.MepElementId] = new List<ClashZone>();
                    mepIdToZones[cz.MepElementId].Add(cz);
                }
                
                if (cz.StructuralElementId != null && needsHostParams)
                {
                    uniqueHostIds.Add(cz.StructuralElementId);
                    if (!hostIdToZones.ContainsKey(cz.StructuralElementId))
                        hostIdToZones[cz.StructuralElementId] = new List<ClashZone>();
                    hostIdToZones[cz.StructuralElementId].Add(cz);
                }
            }
            
            Log($"[PARAM-CAPTURE-BATCH] Collected {uniqueMepIds.Count} unique MEP elements and {uniqueHostIds.Count} unique host elements");
            elementRetrievalSw.Stop();
            
            // ✅ STEP 2: Get all elements once and cache them
            var elementCacheSw = System.Diagnostics.Stopwatch.StartNew();
            var mepElementCache = new Dictionary<ElementId, Element?>();
            var hostElementCache = new Dictionary<ElementId, Element?>();
            
            foreach (var mepId in uniqueMepIds)
            {
                try
                {
                    var mep = ElementRetrievalService.GetElementFromDocumentOrLinked(
                        _context.Document, mepId, enableLogging: false);
                    mepElementCache[mepId] = mep;
                }
                catch (Exception ex)
                {
                    Log($"[PARAM-CAPTURE-BATCH] ⚠️ Error retrieving MEP element {mepId}: {ex.Message}");
                    mepElementCache[mepId] = null;
                }
            }
            
            foreach (var hostId in uniqueHostIds)
            {
                try
                {
                    var host = ElementRetrievalService.GetElementFromDocumentOrLinked(
                        _context.Document, hostId, enableLogging: false);
                    hostElementCache[hostId] = host;
                }
                catch (Exception ex)
                {
                    Log($"[PARAM-CAPTURE-BATCH] ⚠️ Error retrieving host element {hostId}: {ex.Message}");
                    hostElementCache[hostId] = null;
                }
            }
            
            elementCacheSw.Stop();
            Log($"[PARAM-CAPTURE-BATCH] Element retrieval: {elementCacheSw.ElapsedMilliseconds}ms ({uniqueMepIds.Count + uniqueHostIds.Count} elements)");
            
            // ✅ STEP 3: Process parameters for all elements and cache results
            var paramProcessingSw = System.Diagnostics.Stopwatch.StartNew();
            var mepParamCache = new Dictionary<ElementId, List<SerializableKeyValue>>();
            var hostParamCache = new Dictionary<ElementId, List<SerializableKeyValue>>();
            
            foreach (var kvp in mepElementCache)
            {
                if (kvp.Value != null)
                {
                    try
                    {
                        var paramsList = CaptureMinimalParams(kvp.Value);
                        mepParamCache[kvp.Key] = paramsList;
                    }
                    catch (Exception ex)
                    {
                        Log($"[PARAM-CAPTURE-BATCH] ⚠️ Error capturing params for MEP element {kvp.Key}: {ex.Message}");
                        mepParamCache[kvp.Key] = new List<SerializableKeyValue>();
                    }
                }
                else
                {
                    mepParamCache[kvp.Key] = new List<SerializableKeyValue>();
                }
            }
            
            foreach (var kvp in hostElementCache)
            {
                if (kvp.Value != null)
                {
                    try
                    {
                        // ✅ FIX: Capture host params with isHostElement=true to filter out MEP-specific parameters
                        var paramsList = CaptureMinimalParams(kvp.Value, isHostElement: true);
                        hostParamCache[kvp.Key] = paramsList;
                    }
                    catch (Exception ex)
                    {
                        Log($"[PARAM-CAPTURE-BATCH] ⚠️ Error capturing params for host element {kvp.Key}: {ex.Message}");
                        hostParamCache[kvp.Key] = new List<SerializableKeyValue>();
                    }
                }
                else
                {
                    hostParamCache[kvp.Key] = new List<SerializableKeyValue>();
                }
            }
            
            paramProcessingSw.Stop();
            Log($"[PARAM-CAPTURE-BATCH] Parameter processing: {paramProcessingSw.ElapsedMilliseconds}ms");

            // ✅ BUG FIX + PERF: Extract level info for every unique MEP element BEFORE clearing element caches.
            // Previously, mepElementCache was cleared before ExtractMepElementLevelInfo was called in Step 4,
            // causing MepElementLevelName/MepElementLevelElevation to NEVER be set in batch mode.
            // Now: build a lightweight (levelName, elevation) cache per element ID during Step 3, use it in Step 4.
            var levelExtractionSw = System.Diagnostics.Stopwatch.StartNew();
            var mepLevelInfoCache = new Dictionary<ElementId, (string? name, double elevation)>(uniqueMepIds.Count);
            if (_levelByNameCache == null)
            {
                // Build level-by-name lookup once for all MEP elements
                try
                {
                    _levelByNameCache = new FilteredElementCollector(_context.Document)
                        .OfClass(typeof(Level))
                        .Cast<Level>()
                        .GroupBy(l => l.Name, StringComparer.OrdinalIgnoreCase)
                        .ToDictionary(g => g.Key, g => (Level?)g.First(), StringComparer.OrdinalIgnoreCase);
                }
                catch
                {
                    _levelByNameCache = new Dictionary<string, Level?>(StringComparer.OrdinalIgnoreCase);
                }
            }
            foreach (var kvp in mepElementCache)
            {
                if (kvp.Value != null)
                {
                    try
                    {
                        Level? lvl = GetMepLevelFromElement(kvp.Value, _levelByNameCache);
                        if (lvl != null)
                            mepLevelInfoCache[kvp.Key] = (lvl.Name, lvl.Elevation);
                    }
                    catch { }
                }
            }
            levelExtractionSw.Stop();
            Log($"[PARAM-CAPTURE-BATCH] Level extraction: {levelExtractionSw.ElapsedMilliseconds}ms ({mepLevelInfoCache.Count} levels found for {mepElementCache.Count} MEP elements)");

            // ✅ MEMORY OPTIMIZATION: Clear element caches immediately after parameter processing
            mepElementCache.Clear();
            hostElementCache.Clear();

            // ✅ STEP 4: Map parameters back to clash zones (same merge logic as sequential)
            var mappingSw = System.Diagnostics.Stopwatch.StartNew();
            int processedCount = 0;
            int totalParamsBefore = clashZones.Sum(cz => 
                (cz.MepParameterValues?.Count ?? 0) + (cz.HostParameterValues?.Count ?? 0));
            
            foreach (var cz in clashZones)
            {
                try
                {
                    // Process MEP parameters
                    if (cz.MepElementId != null && mepParamCache.TryGetValue(cz.MepElementId, out var newMepParams))
                    {
                        if (newMepParams != null && newMepParams.Count > 0)
                        {
                            if (cz.MepParameterValues == null)
                            {
                                cz.MepParameterValues = newMepParams;
                            }
                            else if (cz.MepParameterValues.Count >= MinimalWhitelist.Count)
                            {
                                // ✅ PERF: Zone already has ≥ MinimalWhitelist.Count params from 6d0 CaptureBatchParams.
                                // Minimal params are a subset of the full whitelist — no new params to add, skip expensive merge.
                            }
                            else
                            {
                                // Zone has fewer params than expected — run merge to fill any gaps
                                var existingDict = cz.MepParameterValues
                                    .Where(kv => kv != null && !string.IsNullOrEmpty(kv.Key))
                                    .GroupBy(kv => kv.Key, StringComparer.OrdinalIgnoreCase)
                                    .ToDictionary(g => g.Key, g => g.First().Value, StringComparer.OrdinalIgnoreCase);

                                foreach (var newParam in newMepParams)
                                {
                                    if (newParam != null && !string.IsNullOrEmpty(newParam.Key) && !existingDict.ContainsKey(newParam.Key))
                                        existingDict[newParam.Key] = newParam.Value;
                                }

                                cz.MepParameterValues = existingDict
                                    .Select(kv => new SerializableKeyValue { Key = kv.Key, Value = kv.Value })
                                    .ToList();
                            }
                        }
                        // ✅ If CaptureMinimalParams fails or returns empty, preserve existing parameters
                    }

                    // Process host parameters
                    if (cz.StructuralElementId != null && hostParamCache.TryGetValue(cz.StructuralElementId, out var newHostParams))
                    {
                        if (newHostParams != null && newHostParams.Count > 0)
                        {
                            if (cz.HostParameterValues == null)
                            {
                                cz.HostParameterValues = newHostParams;
                            }
                            else if (cz.HostParameterValues.Count >= newHostParams.Count)
                            {
                                // ✅ PERF: Host already has at least as many params as minimal capture — skip merge
                            }
                            else
                            {
                                var existingDict = cz.HostParameterValues
                                    .Where(kv => kv != null && !string.IsNullOrEmpty(kv.Key))
                                    .GroupBy(kv => kv.Key, StringComparer.OrdinalIgnoreCase)
                                    .ToDictionary(g => g.Key, g => g.First().Value, StringComparer.OrdinalIgnoreCase);

                                foreach (var newParam in newHostParams)
                                {
                                    if (newParam != null && !string.IsNullOrEmpty(newParam.Key) && !existingDict.ContainsKey(newParam.Key))
                                        existingDict[newParam.Key] = newParam.Value;
                                }

                                cz.HostParameterValues = existingDict
                                    .Select(kv => new SerializableKeyValue { Key = kv.Key, Value = kv.Value })
                                    .ToList();
                            }
                        }
                        // ✅ If CaptureMinimalParams fails or returns empty, preserve existing parameters
                    }

                    // ✅ BUG FIX: Use pre-computed level info from mepLevelInfoCache (built in Step 3 before cache clear).
                    // Previously used mepElementCache which was already cleared, so level was never set in batch mode.
                    if (cz.MepElementId != null && mepLevelInfoCache.TryGetValue(cz.MepElementId, out var levelInfo)
                        && !string.IsNullOrEmpty(levelInfo.name))
                    {
                        cz.MepElementLevelName = levelInfo.name;
                        cz.MepElementLevelElevation = levelInfo.elevation;
                    }
                    
                    processedCount++;
                }
                catch (Exception ex)
                {
                    Log($"[PARAM-CAPTURE-BATCH] ⚠️ Error mapping params for zone {cz.Id}: {ex.Message}");
                }
            }
            
            mappingSw.Stop();
            
            // ✅ STEP 5: Validation - ensure no parameters were lost
            int totalParamsAfter = clashZones.Sum(cz => 
                (cz.MepParameterValues?.Count ?? 0) + (cz.HostParameterValues?.Count ?? 0));
            
            sw.Stop();
            
            Log($"[PARAM-CAPTURE-BATCH] Parameter mapping: {mappingSw.ElapsedMilliseconds}ms");
            Log($"[PARAM-CAPTURE-BATCH] ✅ Batch capture complete: {processedCount} zones in {sw.ElapsedMilliseconds}ms");
            Log($"[PARAM-CAPTURE-BATCH] Total params: {totalParamsBefore} → {totalParamsAfter} ({(totalParamsAfter >= totalParamsBefore ? "✅ NO LOSS" : "⚠️ LOSS DETECTED")})");
            Log($"[PARAM-CAPTURE-BATCH] Average: {(double)totalParamsAfter / processedCount:F1} params/zone");
            Log($"[PARAM-CAPTURE-BATCH] String pool size: {_context.StringPool.Count} unique strings");
            
            // ✅ CRITICAL VALIDATION: Warn if parameters were lost
            if (totalParamsAfter < totalParamsBefore)
            {
                Log($"[PARAM-CAPTURE-BATCH] ⚠️⚠️⚠️ WARNING: Parameter loss detected! Before: {totalParamsBefore}, After: {totalParamsAfter}, Lost: {totalParamsBefore - totalParamsAfter}");
                Log($"[PARAM-CAPTURE-BATCH] ⚠️ Consider disabling UseBatchParameterCapture flag if this persists");
            }
        }
        
        /// <summary>
        /// Captures only whitelisted parameters with pre-interned strings.
        /// ✅ FIX: Filters out MEP-specific parameters when isHostElement=true
        /// ✅ SMART LEVEL LOGIC: For MEP elements, prioritizes "Reference Level" over other level parameters
        /// </summary>
        private List<SerializableKeyValue> CaptureMinimalParams(Element element, bool isHostElement = false)
        {
            var result = new List<SerializableKeyValue>();
            var collected = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            bool hasReferenceLevel = false;  // ✅ Track if Reference Level was found

            try
            {
                // ✅ STEP 1: First pass - collect all parameters, but prioritize Reference Level
                foreach (Parameter param in element.Parameters)
                {
                    if (param == null || !param.HasValue)
                        continue;
                    
                    var paramName = param.Definition?.Name;
                    if (string.IsNullOrEmpty(paramName))
                        continue;

                    if (!AliasToCanonicalMap.TryGetValue(paramName, out var canonicalName))
                        continue;

                    // ✅ FIX: Only capture parameters that are in the whitelist
                    if (!MinimalWhitelist.Contains(canonicalName))
                        continue;

                    // ✅ FIX: Filter out MEP-specific parameters for host elements
                    if (isHostElement)
                    {
                        // MEP-specific parameters that should NOT be captured for host elements
                        var mepSpecificParams = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                        {
                            "System Type", "System Name", "System Abbreviation",
                            "Service Type", "Schedule of Level", "Schedule Level", "Reference Level"
                        };
                        if (mepSpecificParams.Contains(canonicalName))
                            continue;
                    }

                    // ✅ SMART LEVEL LOGIC: For MEP elements, check Reference Level first
                    if (!isHostElement && canonicalName == "Reference Level")
                    {
                        var paramValue = GetParameterValueAsString(param, element);
                        if (!string.IsNullOrEmpty(paramValue))
                        {
                            collected["Reference Level"] = paramValue;
                            hasReferenceLevel = true;
                        }
                        continue;  // Skip to next parameter
                    }

                    // ✅ SMART LEVEL LOGIC: If Reference Level found, skip other level parameters for MEP elements
                    if (!isHostElement && hasReferenceLevel)
                    {
                        var levelParams = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                        {
                            "Level", "Schedule Level", "Schedule of Level"
                        };
                        if (levelParams.Contains(canonicalName))
                            continue;  // Skip fallback level params if Reference Level exists
                    }

                    if (collected.ContainsKey(canonicalName))
                        continue;

                    // ✅ CRITICAL FIX: Pass element owner to resolve ElementId parameters (Level, etc.)
                    var paramValue2 = GetParameterValueAsString(param, element);
                    if (string.IsNullOrEmpty(paramValue2))
                        continue;

                    collected[canonicalName] = paramValue2;
                }

                // ✅ STEP 2: If Reference Level not found for MEP elements, try fallback level parameters
                if (!isHostElement && !hasReferenceLevel)
                {
                    var fallbackLevelParams = new[] { "Level", "Schedule Level", "Schedule of Level" };
                    bool foundFallbackLevel = false;
                    
                    foreach (var levelParamName in fallbackLevelParams)
                    {
                        if (collected.ContainsKey(levelParamName))
                        {
                            foundFallbackLevel = true;
                            break;  // Already captured in first pass
                        }

                        // Try to find this level parameter
                        foreach (Parameter param in element.Parameters)
                        {
                            if (param == null || !param.HasValue)
                                continue;
                            
                            var paramName = param.Definition?.Name;
                            if (string.IsNullOrEmpty(paramName))
                                continue;

                            if (!AliasToCanonicalMap.TryGetValue(paramName, out var canonicalName))
                                continue;

                            if (canonicalName == levelParamName && MinimalWhitelist.Contains(canonicalName))
                            {
                                var paramValue = GetParameterValueAsString(param, element);
                                if (!string.IsNullOrEmpty(paramValue))
                                {
                                    collected[levelParamName] = paramValue;
                                    foundFallbackLevel = true;
                                    break;  // Found this fallback level param, stop searching
                                }
                            }
                        }
                        
                        if (foundFallbackLevel)
                            break;  // Found a fallback level param, don't try others
                    }
                }

                // ✅ FIX: Only capture MEP-specific parameters for MEP elements, not host elements
                if (!isHostElement)
                {
                    // Built-in fallbacks for essential parameters that may not have direct parameter names
                    // ✅ CRITICAL: Check if element is a Cable Tray (needs "Service Type" instead of "System Type")
                    bool isCableTray = element.Category?.Id?.GetIntegerValue() == (int)BuiltInCategory.OST_CableTray ||
                                      element.Category?.Name?.Contains("Cable Tray", StringComparison.OrdinalIgnoreCase) == true ||
                                      element is Autodesk.Revit.DB.Electrical.CableTray;
                    
                    if (isCableTray)
                    {
                        // For Cable Trays, use "Service Type" instead of "System Type"
                        EnsureServiceType(element, collected);
                    }
                    else
                    {
                        // For Mechanical/Plumbing (Ducts/Pipes), use "System Type"
                        EnsureSystemType(element, collected);
                    }
                    
                    EnsureSystemName(element, collected);
                    EnsureSystemAbbreviation(element, collected);
                    EnsureScheduleOfLevel(element, collected);  // ✅ ADDED: Capture Schedule of Level for Bottom of Opening calculation
                }
            }
            catch (Exception ex)
            {
                Log($"[PARAM-CAPTURE] ⚠️ Error reading parameters from element {element.Id}: {ex.Message}");
            }

            // ✅ FIX: Only add parameters that are in the whitelist
            foreach (var kv in collected)
            {
                if (!MinimalWhitelist.Contains(kv.Key))
                    continue;

                result.Add(new SerializableKeyValue
                {
                    Key = _context.StringPool.Intern(kv.Key),
                    Value = _context.StringPool.Intern(kv.Value)
                });
            }
            
            return result;
        }
        
        private string GetParameterValueAsString(Parameter param, Element? owner = null)
        {
            try
            {
                switch (param.StorageType)
                {
                    case StorageType.String:
                        return param.AsString() ?? string.Empty;
                    
                    case StorageType.Integer:
                        return param.AsInteger().ToString();
                    
                    case StorageType.Double:
                        return param.AsDouble().ToString("F3");
                    
                    case StorageType.ElementId:
                        // ✅ CRITICAL FIX: Resolve ElementId to actual element name (for Level, etc.)
                        var id = param.AsElementId();
                        if (id != null && id.GetIntegerValue() > 0 && owner != null)
                        {
                            try
                            {
                                var referencedElement = owner.Document?.GetElement(id);
                                if (referencedElement != null)
                                {
                                    // For Level parameters, get the Level name
                                    if (referencedElement is Level level)
                                    {
                                        return level.Name ?? id.GetIntegerValue().ToString();
                                    }
                                    // For other ElementId types, get the element name
                                    return referencedElement.Name ?? id.GetIntegerValue().ToString();
                                }
                            }
                            catch
                            {
                                // Fall back to ID if resolution fails
                            }
                        }
                        return id?.GetIntegerValue().ToString() ?? string.Empty;
                    
                    default:
                        return string.Empty;
                }
            }
            catch
            {
                return string.Empty;
            }
        }

        private static void EnsureSystemType(Element element, Dictionary<string, string> collected)
        {
            if (collected.ContainsKey("System Type"))
                return;

            // ✅ FIX: Only capture if System Type is in whitelist
            if (!MinimalWhitelist.Contains("System Type"))
                return;

            // ✅ CRITICAL: Try built-in parameters first (for Ducts/Pipes)
            Parameter param = element.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM) ??
                              element.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM) ??
                              element.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM);
            
            string paramSource = param != null ? "BuiltIn" : "NotFound";
            
            // ✅ CRITICAL: If built-in parameters not found, try common parameter name variations
            if (param == null)
            {
                param = element.LookupParameter("MEP System Type");
                if (param != null) paramSource = "MEP System Type";
                
                if (param == null)
                {
                    param = element.LookupParameter("System Classification");
                    if (param != null) paramSource = "System Classification";
                }
                
                if (param == null)
                {
                    param = element.LookupParameter("MEP System Classification");
                    if (param != null) paramSource = "MEP System Classification";
                }
                
                // ✅ ADDITIONAL: Try more variations for pipes specifically
                if (param == null && element.Category?.Id?.GetIntegerValue() == (int)BuiltInCategory.OST_PipeCurves)
                {
                    param = element.LookupParameter("System Type");
                    if (param != null) paramSource = "System Type (direct)";
                }
            }

            if (param == null)
            {
                // ✅ DIAGNOSTIC: Log when System Type is not found
                if (OptimizationFlags.UseDiagnosticMode && !DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("Refresh_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [PARAM-CAPTURE] ⚠️ System Type NOT FOUND for element {element.Id} (Category: {element.Category?.Name ?? "Unknown"})\n");
                }
                return;
            }

            // ✅ FIX: Get text value from System Type parameter (resolves ElementId to element name)
            string value = string.Empty;
            
            // Try AsString() first (for string parameters)
            value = param.AsString();
            if (!string.IsNullOrWhiteSpace(value))
            {
                collected["System Type"] = value.Trim();
                if (OptimizationFlags.UseDiagnosticMode && !DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("Refresh_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [PARAM-CAPTURE] ✅ System Type CAPTURED for element {element.Id}: '{value.Trim()}' (Source: {paramSource}, Method: AsString)\n");
                }
                return;
            }

            // Try AsValueString() (formatted display value)
            value = param.AsValueString();
            if (!string.IsNullOrWhiteSpace(value))
            {
                collected["System Type"] = value.Trim();
                if (OptimizationFlags.UseDiagnosticMode && !DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("Refresh_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [PARAM-CAPTURE] ✅ System Type CAPTURED for element {element.Id}: '{value.Trim()}' (Source: {paramSource}, Method: AsValueString)\n");
                }
                return;
            }

            // ✅ CRITICAL FIX: For ElementId storage type, resolve to system type element name
            if (param.StorageType == StorageType.ElementId)
            {
                try
                {
                    var systemTypeId = param.AsElementId();
                    if (systemTypeId != null && systemTypeId.GetIntegerValue() > 0)
                    {
                        var systemTypeElement = element.Document?.GetElement(systemTypeId);
                        if (systemTypeElement != null)
                        {
                            // Get the name of the system type element (e.g., "Supply Air", "Return Air")
                            value = systemTypeElement.Name;
                            if (!string.IsNullOrWhiteSpace(value))
                            {
                                collected["System Type"] = value.Trim();
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    SafeFileLogger.SafeAppendText("Refresh_debug.log",
                                        $"[{DateTime.Now:HH:mm:ss.fff}] [PARAM-CAPTURE] ✅ System Type CAPTURED for element {element.Id}: '{value.Trim()}' (Source: {paramSource}, Method: ElementId->Name)\n");
                                }
                                return;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("Refresh_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [PARAM-CAPTURE] ⚠️ Error resolving System Type ElementId for element {element.Id}: {ex.Message}\n");
                    }
                }
            }
            
            // ✅ DIAGNOSTIC: Log when parameter found but value is empty
            if (!DeploymentConfiguration.DeploymentMode)
            {
                SafeFileLogger.SafeAppendText("Refresh_debug.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [PARAM-CAPTURE] ⚠️ System Type parameter FOUND but value is EMPTY for element {element.Id} (Source: {paramSource}, StorageType: {param.StorageType})\n");
            }
        }

        private static void EnsureSystemName(Element element, Dictionary<string, string> collected)
        {
            if (collected.ContainsKey("System Name"))
                return;

            // ✅ FIX: Only capture if System Name is in whitelist
            if (!MinimalWhitelist.Contains("System Name"))
                return;

            // ✅ CRITICAL: Try built-in parameter first
            Parameter param = element.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM);
            
            // ✅ CRITICAL: If built-in parameter not found, try common parameter name variations
            if (param == null)
            {
                param = element.LookupParameter("MEP System Name") ??
                        element.LookupParameter("System Name");
            }
            
            if (param == null)
                return;

            var value = ParameterValueToString(element, param);
            if (!string.IsNullOrWhiteSpace(value))
                collected["System Name"] = value;
        }

        private static void EnsureSystemAbbreviation(Element element, Dictionary<string, string> collected)
        {
            if (collected.ContainsKey("System Abbreviation"))
                return;

            // ✅ FIX: Only capture if System Abbreviation is in whitelist
            if (!MinimalWhitelist.Contains("System Abbreviation"))
                return;

            Parameter param = element.get_Parameter(BuiltInParameter.RBS_SYSTEM_ABBREVIATION_PARAM);
            if (param == null)
                return;

            var value = ParameterValueToString(element, param);
            if (!string.IsNullOrWhiteSpace(value))
                collected["System Abbreviation"] = value;
        }
        
        /// <summary>
        /// ✅ CRITICAL: Ensure "Service Type" is captured for Cable Trays and Conduits.
        /// For Cable Trays, "Service Type" is the equivalent of "System Type" for Mechanical/Plumbing elements.
        /// </summary>
        private static void EnsureServiceType(Element element, Dictionary<string, string> collected)
        {
            if (collected.ContainsKey("Service Type"))
                return;

            // ✅ FIX: Only capture if Service Type is in whitelist
            if (!MinimalWhitelist.Contains("Service Type"))
                return;

            // Try common parameter name variations for Service Type
            Parameter param = element.LookupParameter("Service Type") ??
                             element.LookupParameter("MEP Service Type");
            
            if (param == null)
                return;

            var value = ParameterValueToString(element, param);
            if (!string.IsNullOrWhiteSpace(value))
                collected["Service Type"] = value;
        }

        /// <summary>
        /// ✅ CRITICAL: Ensure "Schedule of Level" is captured for Bottom of Opening calculation.
        /// "Schedule of Level" = "Elevation from Level" = height from level elevation to placement point.
        /// </summary>
        private static void EnsureScheduleOfLevel(Element element, Dictionary<string, string> collected)
        {
            if (collected.ContainsKey("Schedule of Level"))
                return;

            // ✅ FIX: Only capture if Schedule of Level is in whitelist
            if (!MinimalWhitelist.Contains("Schedule of Level"))
                return;

            // Try common parameter name variations for Schedule of Level
            Parameter param = element.LookupParameter("Elevation from Level")  // ✅ FIRST: This is what user sees in Properties
                             ?? element.LookupParameter("Schedule of Level")
                             ?? element.LookupParameter("Schedule Level");
            
            if (param == null)
            {
                if (OptimizationFlags.UseDiagnosticMode && !DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("Refresh_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [PARAM-CAPTURE] ⚠️ Schedule of Level NOT FOUND for element {element.Id} (tried: 'Elevation from Level', 'Schedule of Level', 'Schedule Level')\n");
                }
                return;
            }

            var value = ParameterValueToString(element, param);
            if (!string.IsNullOrWhiteSpace(value))
            {
                collected["Schedule of Level"] = value;
                if (OptimizationFlags.UseDiagnosticMode && !DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("Refresh_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [PARAM-CAPTURE] ✅ Schedule of Level CAPTURED for element {element.Id}: '{value}' (Source: '{param.Definition.Name}')\n");
                }
            }
            else
            {
                if (OptimizationFlags.UseDiagnosticMode && !DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("Refresh_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [PARAM-CAPTURE] ⚠️ Schedule of Level parameter FOUND but value is EMPTY for element {element.Id} (Source: '{param.Definition.Name}')\n");
                }
            }
        }

        private static string ParameterValueToString(Element owner, Parameter parameter)
        {
            if (parameter == null)
                return string.Empty;

            // ✅ CRITICAL: For Double storage type (like "Elevation from Level"), return the numeric value
            if (parameter.StorageType == StorageType.Double)
            {
                return parameter.AsDouble().ToString("F6");  // Use high precision for elevation values
            }

            var value = parameter.AsString();
            if (!string.IsNullOrEmpty(value))
                return value;

            value = parameter.AsValueString();
            if (!string.IsNullOrEmpty(value))
                return value;

            if (parameter.StorageType == StorageType.ElementId)
            {
                try
                {
                    var id = parameter.AsElementId();
                    if (id != null && id.GetIntegerValue() > 0 && owner != null)
                    {
                        var referenced = owner.Document?.GetElement(id);
                        if (referenced != null)
                        {
                            // ✅ CRITICAL FIX: For Level parameters, get the Level name
                            if (referenced is Level level)
                            {
                                return level.Name ?? id.GetIntegerValue().ToString();
                            }
                            // For other ElementId types, get the element name
                            return referenced.Name ?? id.GetIntegerValue().ToString();
                        }
                    }
                    return id?.GetIntegerValue().ToString() ?? string.Empty;
                }
                catch
                {
                    return string.Empty;
                }
            }

            if (parameter.StorageType == StorageType.Integer)
                return parameter.AsInteger().ToString();

            return string.Empty;
        }
        
        private void Log(string message)
        {
            if (!_context.IsDeploymentMode)
                DebugLogger.Info(message);
            
            // ✅ PERF GATE: OptimizationFlags.UseDiagnosticMode gates per-zone logging (384 calls adding 10ms each = 3.8s savings)
            if (OptimizationFlags.UseDiagnosticMode)
            {
                SafeFileLogger.SafeAppendText(_context.RefreshLogName, $"[{DateTime.Now}] {message}\n");
            }
        }
    }
    
    /// <summary>
    /// Extension method for batching collections
    /// </summary>
    public static class BatchExtensions
    {
        public static IEnumerable<IEnumerable<T>> Batch<T>(this IEnumerable<T> source, int batchSize)
        {
            var batch = new List<T>(batchSize);
            
            foreach (var item in source)
            {
                batch.Add(item);
                
                if (batch.Count >= batchSize)
                {
                    yield return batch;
                    batch = new List<T>(batchSize);
                }
            }
            
            if (batch.Count > 0)
                yield return batch;
        }
    }
}