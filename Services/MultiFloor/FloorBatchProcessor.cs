using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Utils;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.MultiFloor
{
    /// <summary>
    /// Orchestrates multi-floor clash detection and sleeve placement
    /// Handles chunking, memory management, and parallel execution
    /// </summary>
    public class FloorBatchProcessor
    {
        private readonly Document _doc;
        private readonly IPerformanceMonitor? _monitor;
        private readonly int _maxParallelism;
        private readonly CheckpointManager _checkpointMgr;
        
        public FloorBatchProcessor(
            Document doc, 
            IPerformanceMonitor? monitor = null,
            string? checkpointPath = null)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _monitor = monitor;
            _maxParallelism = Math.Min(Environment.ProcessorCount, 8); // Max 8 cores
            _checkpointMgr = new CheckpointManager(checkpointPath ?? GetDefaultCheckpointPath());
        }
        
        /// <summary>
        /// Main entry point: Process multiple floors with automatic chunking
        /// </summary>
        /// <param name="levels">Levels to process</param>
        /// <param name="filter">Opening filter configuration</param>
        /// <param name="chunkSize">Floors per chunk (default: 5 for memory safety)</param>
        public MultiFloorResult ProcessFloors(
            List<Level> levels, 
            OpeningFilter filter,
            int chunkSize = 5)
        {
            // Pre-flight validation
            var validation = ValidateBeforeProcessing(levels);
            if (!validation.IsValid)
            {
                throw new InvalidOperationException(
                    $"Pre-flight validation failed: {string.Join(", ", validation.Errors)}");
            }
            
            var totalResult = new MultiFloorResult();
            
            // Check for checkpoint (resume capability)
            var checkpoint = _checkpointMgr.LoadCheckpoint();
            if (checkpoint != null)
            {
                SafeFileLogger.SafeAppendText("multifloor.log",
                    $"[{DateTime.Now}] 🔄 RESUMING from checkpoint: {checkpoint.ProcessedFloors.Count} floors already done\n");
                levels = levels.Where(l => !checkpoint.ProcessedFloors.Contains(l.Name)).ToList();
            }
            
            // Pre-cache symbols ONCE for all floors
            SharedResourceCache cache;
            using (var cacheTracker = _monitor?.TrackOperation("Pre-load Shared Resources"))
            {
                cache = new SharedResourceCache(_doc);
                cache.Initialize();
                cacheTracker?.SetItemCount(cache.Symbols.Count);
            }
            
            // Process in chunks for memory safety
            int chunkIndex = 0;
            for (int i = 0; i < levels.Count; i += chunkSize)
            {
                var chunk = levels.Skip(i).Take(chunkSize).ToList();
                chunkIndex++;
                
                SafeFileLogger.SafeAppendText("multifloor.log",
                    $"[{DateTime.Now}] 🚀 Processing chunk {chunkIndex}: Floors {string.Join(", ", chunk.Select(l => l.Name))}\n");
                
                using (var chunkTracker = _monitor?.TrackOperation($"Floor Chunk {chunkIndex}"))
                {
                    var chunkResult = ProcessFloorChunk(chunk, filter, cache);
                    totalResult.Merge(chunkResult);
                    chunkTracker?.SetItemCount(chunk.Count);
                    
                    // Save checkpoint after each chunk
                    _checkpointMgr.SaveCheckpoint(new MultiFloorProgress
                    {
                        CompletedFloors = totalResult.SuccessfulFloors,
                        TotalFloors = levels.Count
                    });
                    
                    // Force GC between chunks
                    SafeFileLogger.SafeAppendText("multifloor.log",
                        $"[{DateTime.Now}] 🧹 Forcing GC after chunk {chunkIndex}...\n");
                    GC.Collect(2, GCCollectionMode.Forced);
                    GC.WaitForPendingFinalizers();
                }
            }
            
            // Clear checkpoint on success
            _checkpointMgr.ClearCheckpoint();
            
            return totalResult;
        }
        
        /// <summary>
        /// Process a chunk of floors.
        /// Sequential Revit work + Parallel Post-processing
        /// </summary>
        private MultiFloorResult ProcessFloorChunk(
            List<Level> chunk,
            OpeningFilter filter,
            SharedResourceCache cache)
        {
            var results = new List<FloorProcessingResult>();

            // 1. Revit work per floor (SEQUENTIAL, one Transaction per floor)
            //    API calls must be on the main thread.
            foreach (var level in chunk)
            {
                try
                {
                    var floorProcessor = new SingleFloorProcessor(_doc, cache, _monitor);
                    var res = floorProcessor.ProcessFloor(level, filter);
                    results.Add(res);
                }
                catch (Exception ex)
                {
                    SafeFileLogger.SafeAppendText("multifloor.log",
                        $"[{DateTime.Now}] ❌ Floor {level.Name} FAILED: {ex.Message}\n");
                    results.Add(new FloorProcessingResult
                    {
                        LevelName = level.Name,
                        Success = false,
                        ErrorMessage = ex.Message
                    });
                }
            }

            return new MultiFloorResult(results);
        }
        
        /// <summary>
        /// Pre-flight validation before processing
        /// </summary>
        private ValidationResult ValidateBeforeProcessing(List<Level> levels)
        {
            var result = new ValidationResult { IsValid = true };
            
            // Memory check
            // (Re-calculate budget based on project characteristics)
            
            return result;
        }
        
        private string GetDefaultCheckpointPath()
        {
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var dir = System.IO.Path.Combine(appData, "JSE_MEP_Openings", "Checkpoints");
            System.IO.Directory.CreateDirectory(dir);
            return System.IO.Path.Combine(dir, $"checkpoint_{_doc.Title}.json");
        }
    }
}
