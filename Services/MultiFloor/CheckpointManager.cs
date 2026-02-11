using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using JSE_RevitAddin_MEP_OPENINGS.Utils;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.MultiFloor
{
    /// <summary>
    /// Manages checkpoints for crash recovery and resume
    /// </summary>
    public class CheckpointManager
    {
        private readonly string _checkpointPath;
        
        public CheckpointManager(string checkpointPath)
        {
            _checkpointPath = checkpointPath;
        }
        
        public void SaveCheckpoint(MultiFloorProgress progress)
        {
            try
            {
                var checkpoint = new Checkpoint
                {
                    Timestamp = DateTime.Now,
                    ProcessedFloors = progress.CompletedFloors,
                    TotalFloors = progress.TotalFloors
                };
                
                var json = JsonSerializer.Serialize(checkpoint, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(_checkpointPath, json);
                
                SafeFileLogger.SafeAppendText("checkpoint.log",
                    $"[{DateTime.Now}] 💾 Checkpoint saved: {progress.CompletedFloors.Count}/{progress.TotalFloors} floors\n");
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("checkpoint.log",
                    $"[{DateTime.Now}] ⚠️ Failed to save checkpoint: {ex.Message}\n");
            }
        }
        
        public Checkpoint? LoadCheckpoint()
        {
            if (!File.Exists(_checkpointPath))
                return null;
            
            try
            {
                var json = File.ReadAllText(_checkpointPath);
                var checkpoint = JsonSerializer.Deserialize<Checkpoint>(json);
                
                if (checkpoint == null) return null;
                
                // Check if checkpoint is recent (within last 24 hours)
                if ((DateTime.Now - checkpoint.Timestamp).TotalHours > 24)
                {
                    SafeFileLogger.SafeAppendText("checkpoint.log",
                        $"[{DateTime.Now}] ⚠️ Checkpoint too old ({checkpoint.Timestamp}), ignoring\n");
                    return null;
                }
                
                return checkpoint;
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("checkpoint.log",
                    $"[{DateTime.Now}] ⚠️ Failed to load checkpoint: {ex.Message}\n");
                return null;
            }
        }
        
        public void ClearCheckpoint()
        {
            try
            {
                if (File.Exists(_checkpointPath))
                {
                    File.Delete(_checkpointPath);
                    SafeFileLogger.SafeAppendText("checkpoint.log",
                        $"[{DateTime.Now}] 🗑️ Checkpoint cleared\n");
                }
            }
            catch { }
        }
    }
}
