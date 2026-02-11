using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Utils;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Proximity
{
    /// <summary>
    /// Service for marking zones eligible for clustering based on proximity analysis.
    /// Executes the "Proximity Check" step of the clustering workflow (Step 3 of methodology).
    /// </summary>
    public class ClusterProximityMarker
    {
        private readonly IClashZoneRepository _repository;
        private readonly double _proximityTolerance;
        private readonly Action<string> _logger;

        public ClusterProximityMarker(IClashZoneRepository repository, double proximityTolerance, Action<string> logger = null)
        {
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
            _proximityTolerance = proximityTolerance;
            _logger = logger ?? (msg => { });
        }

        /// <summary>
        /// Marks zones for clustering based on proximity analysis.
        ///
        /// Workflow:
        /// 1. Load zones with ReadyForPlacementFlag=1 and IsResolvedFlag=1
        /// 2. For each zone, check if any other zone is within proximity tolerance
        /// 3. Set MarkedForClusterProcess=1 if neighbors exist, 0 if isolated
        /// 4. Batch update database
        /// </summary>
        public int MarkZonesForClustering(Document doc, string targetCategory = null)
        {
            int markedCount = 0;

            try
            {
                _logger("[PROXIMITY-MARKER] 🔍 Starting proximity analysis for clustering eligibility...");

                // ✅ DIAGNOSTIC: Log entry parameters
                SafeFileLogger.SafeAppendText("flag_workflow.log",
                    $"[{DateTime.Now:HH:mm:ss}]   MarkZonesForClustering called with targetCategory={targetCategory ?? "NULL"}\n");

                // Step 1: Load zones ready for proximity check
                var zones = _repository.GetZonesReadyForProximityCheck(targetCategory);

                SafeFileLogger.SafeAppendText("flag_workflow.log",
                    $"[{DateTime.Now:HH:mm:ss}]   GetZonesReadyForProximityCheck returned {zones.Count} zones\n");

                if (zones.Count == 0)
                {
                    _logger("[PROXIMITY-MARKER] ℹ️ No zones ready for proximity checking.");
                    SafeFileLogger.SafeAppendText("flag_workflow.log",
                        $"[{DateTime.Now:HH:mm:ss}]   ⚠️ No zones returned - proximity check cannot proceed\n");
                    SafeFileLogger.SafeAppendText("flag_workflow.log",
                        $"[{DateTime.Now:HH:mm:ss}]   HINT: Check if targetCategory='{targetCategory}' matches database MepCategory values\n");
                    SafeFileLogger.SafeAppendText("flag_workflow.log",
                        $"[{DateTime.Now:HH:mm:ss}]   Common mismatch: 'CableTrays' vs 'Cable Trays' (space)\n");
                    return 0;
                }

                _logger($"[PROXIMITY-MARKER] Analyzing {zones.Count} zones for proximity...");

                // ✅ DIAGNOSTIC: Log before proximity analysis
                SafeFileLogger.SafeAppendText("flag_workflow.log",
                    $"[{DateTime.Now:HH:mm:ss}]   Starting proximity analysis for {zones.Count} zones...\n");

                // Step 2: Run parallel proximity check
                var updates = new ConcurrentBag<(Guid zoneId, bool hasProximity)>();
                var proximityErrors = new ConcurrentBag<string>();

                System.Threading.Tasks.Parallel.ForEach(zones, new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount },
                    zone =>
                    {
                        try
                        {
                            bool hasNearbyZone = CheckProximityToOtherZones(zone, zones);
                            updates.Add((zone.Id, hasNearbyZone));

                            lock (updates) // Thread-safe increment
                            {
                                if (hasNearbyZone)
                                {
                                    markedCount++;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            var errorMsg = $"Zone {zone.Id.ToString().Substring(0, 8)}: {ex.Message}";
                            proximityErrors.Add(errorMsg);
                            _logger($"[PROXIMITY-MARKER] ⚠️ Error checking proximity for zone {zone.Id}: {ex.Message}");
                            SafeFileLogger.SafeAppendText("flag_workflow.log",
                                $"[{DateTime.Now:HH:mm:ss}]   ⚠️ PROXIMITY ERROR: {errorMsg}\n");
                            SafeFileLogger.SafeAppendText("flag_workflow.log",
                                $"[{DateTime.Now:HH:mm:ss}]      Stack: {ex.StackTrace}\n");
                            // Mark as not having proximity if error occurs
                            updates.Add((zone.Id, false));
                        }
                    });

                // ✅ DIAGNOSTIC: Log after proximity analysis
                SafeFileLogger.SafeAppendText("flag_workflow.log",
                    $"[{DateTime.Now:HH:mm:ss}]   Proximity analysis completed. Updates collected: {updates.Count}\n");
                if (proximityErrors.Count > 0)
                {
                    SafeFileLogger.SafeAppendText("flag_workflow.log",
                        $"[{DateTime.Now:HH:mm:ss}]   ⚠️ {proximityErrors.Count} zones had errors during proximity check\n");
                }

                // Step 3: Batch update MarkedForClusterProcess flags
                if (updates.Count > 0)
                {
                    var updateList = updates.ToList();

                    // ✅ STEP-BY-STEP DIAGNOSTIC LOG
                    SafeFileLogger.SafeAppendText("flag_workflow.log",
                        $"[{DateTime.Now:HH:mm:ss}] STEP 4: PROXIMITY CHECK COMPLETED\n");
                    SafeFileLogger.SafeAppendText("flag_workflow.log",
                        $"[{DateTime.Now:HH:mm:ss}]   Analyzed {zones.Count} zones\n");
                    SafeFileLogger.SafeAppendText("flag_workflow.log",
                        $"[{DateTime.Now:HH:mm:ss}]   Found {markedCount} zones WITH neighbors (will set flag=1)\n");
                    SafeFileLogger.SafeAppendText("flag_workflow.log",
                        $"[{DateTime.Now:HH:mm:ss}]   Found {zones.Count - markedCount} isolated zones (will set flag=0)\n");

                    _repository.BatchUpdateMarkedForClusterProcess(updateList);

                    SafeFileLogger.SafeAppendText("flag_workflow.log",
                        $"[{DateTime.Now:HH:mm:ss}]   ✅ BatchUpdateMarkedForClusterProcess completed\n");

                    _logger($"[PROXIMITY-MARKER] ✅ Marked {markedCount}/{zones.Count} zones for clustering (have neighbors)");
                }

                return markedCount;
            }
            catch (Exception ex)
            {
                _logger($"[PROXIMITY-MARKER] ❌ ERROR: {ex.Message}");
                SafeFileLogger.SafeAppendText("proximity_error.log", $"[{DateTime.Now:HH:mm:ss}] MarkZonesForClustering Failed: {ex.Message}\n{ex.StackTrace}\n");

                // ✅ DIAGNOSTIC: Log to flag_workflow.log as well
                SafeFileLogger.SafeAppendText("flag_workflow.log",
                    $"[{DateTime.Now:HH:mm:ss}] ❌❌❌ CRITICAL: PROXIMITY CHECK FAILED ❌❌❌\n");
                SafeFileLogger.SafeAppendText("flag_workflow.log",
                    $"[{DateTime.Now:HH:mm:ss}]   Error: {ex.Message}\n");
                SafeFileLogger.SafeAppendText("flag_workflow.log",
                    $"[{DateTime.Now:HH:mm:ss}]   Stack: {ex.StackTrace}\n");

                return 0;
            }
        }

        /// <summary>
        /// Checks if a zone has any other zones within the proximity tolerance.
        /// Uses appropriate proximity checker based on sleeve shapes.
        /// </summary>
        private bool CheckProximityToOtherZones(ClashZone targetZone, List<ClashZone> allZones)
        {
            if (allZones.Count <= 1)
                return false;

            // For each other zone, check if it's within proximity
            foreach (var otherZone in allZones)
            {
                // Skip self-comparison
                if (targetZone.Id == otherZone.Id)
                    continue;

                // Skip different MEP element categories
                if (!string.IsNullOrEmpty(targetZone.MepElementCategory) && !string.IsNullOrEmpty(otherZone.MepElementCategory))
                {
                    if (targetZone.MepElementCategory != otherZone.MepElementCategory)
                        continue;
                }

                // Create dynamic sleeve objects for checker compatibility
                dynamic sleeve1 = new { ClashZone = targetZone };
                dynamic sleeve2 = new { ClashZone = otherZone };

                try
                {
                    // Get appropriate proximity checker using factory
                    var checker = ProximityCheckerFactory.CreateChecker(sleeve1, sleeve2, 0.0, false);
                    if (checker.CheckProximity(sleeve1, sleeve2, _proximityTolerance))
                    {
                        return true; // Found a nearby zone
                    }
                }
                catch (Exception ex)
                {
                    _logger?.Invoke($"[PROXIMITY-MARKER] ⚠️ Error during proximity check between zones: {ex.Message}");
                    // If checker fails, assume not proximate and continue
                    continue;
                }
            }

            return false; // No nearby zones found
        }
    }
}
