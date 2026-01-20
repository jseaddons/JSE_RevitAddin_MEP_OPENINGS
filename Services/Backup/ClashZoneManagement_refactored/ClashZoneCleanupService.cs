using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor;
using JSE_RevitAddin_MEP_OPENINGS.Services.Logging;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.ClashZoneManagement
{
    /// <summary>
    /// Team G: SOLID-compliant clash zone cleanup service.
    /// 
    /// This service handles cleanup of invalid clash zones and duplicates.
    /// SOLID: Single Responsibility - cleanup only.
    /// 
    /// ✅ PRESERVES ALL LOGIC:
    /// - Step 1: Remove invalid clash zones (null elements)
    /// - Step 2: Remove duplicates (keep first occurrence of each MEP+Structural pair)
    /// - Fail-safe error handling (continue on error, log and skip)
    /// </summary>
    public class ClashZoneCleanupService : IClashZoneCleanupService
    {
        private readonly ILogger _logger;
        
        /// <summary>
        /// Creates a new clash zone cleanup service.
        /// </summary>
        /// <param name="logger">Optional logger for tracking operations</param>
        public ClashZoneCleanupService(ILogger logger = null)
        {
            _logger = logger ?? LoggerAdapter.Default;
        }
        
        /// <summary>
        /// Remove invalid clash zones (null elements) and duplicates.
        /// </summary>
        public int CleanupInvalidClashZones(
            List<ClashZone> clashZones,
            Document document)
        {
            if (clashZones == null || clashZones.Count == 0)
            {
                _logger.Info("No clash zones to clean up", "ClashZoneCleanup");
                return 0;
            }
            
            if (document == null)
            {
                _logger.Warning("Document is null, cannot clean up clash zones", "ClashZoneCleanup");
                return 0;
            }
            
            var originalCount = clashZones.Count;
            _logger.Info($"Starting cleanup of {originalCount} clash zones", "ClashZoneCleanup");
            
            // ✅ STEP 1: Remove invalid clash zones (null elements)
            var validClashZones = new List<ClashZone>();
            var invalidCount = 0;
            
            foreach (var clashZone in clashZones)
            {
                try
                {
                    // ✅ PRESERVE: Check if elements exist in document
                    var mepElement = (clashZone.MepElementId != null && clashZone.MepElementId.IntegerValue != -1) 
                        ? document.GetElement(clashZone.MepElementId) 
                        : null;
                    var structuralElement = (clashZone.StructuralElementId != null && clashZone.StructuralElementId.IntegerValue != -1) 
                        ? document.GetElement(clashZone.StructuralElementId) 
                        : null;
                    
                    if (mepElement != null && structuralElement != null)
                    {
                        validClashZones.Add(clashZone);
                    }
                    else
                    {
                        invalidCount++;
                        _logger.Debug($"Removing invalid clash zone {clashZone.Id} - MEP: {mepElement != null}, Structural: {structuralElement != null}", "ClashZoneCleanup");
                    }
                }
                catch (Exception ex)
                {
                    // ✅ PRESERVE: Fail-safe error handling - continue on error
                    invalidCount++;
                    _logger.Warning($"Removing invalid clash zone {clashZone.Id} due to error: {ex.Message}", "ClashZoneCleanup");
                }
            }
            
            // ✅ STEP 2: Remove duplicates (keep only the first occurrence of each MEP+Structural pair)
            var uniqueClashZones = new List<ClashZone>();
            var seenPairs = new HashSet<(ElementId, ElementId)>();
            var duplicateCount = 0;
            
            foreach (var clashZone in validClashZones)
            {
                var pair = (clashZone.MepElementId, clashZone.StructuralElementId);
                
                if (seenPairs.Add(pair))
                {
                    uniqueClashZones.Add(clashZone);
                }
                else
                {
                    duplicateCount++;
                    _logger.Debug($"Removing duplicate clash zone {clashZone.Id} - MEP: {clashZone.MepElementId}, Structural: {clashZone.StructuralElementId}", "ClashZoneCleanup");
                }
            }
            
            // ✅ UPDATE: Remove invalid and duplicate clash zones from input list
            clashZones.Clear();
            clashZones.AddRange(uniqueClashZones);
            
            var finalCount = clashZones.Count;
            var totalRemoved = originalCount - finalCount;
            
            _logger.Info($"Cleanup complete: {originalCount} → {finalCount} clash zones (removed {totalRemoved}: {invalidCount} invalid + {duplicateCount} duplicates)", "ClashZoneCleanup");
            
            return totalRemoved;
        }
    }
}

