using System;
using System.Collections.Generic;
using System.Linq;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.MultiFloor
{
    public class MultiFloorResult
    {
        public List<string> SuccessfulFloors { get; set; } = new List<string>();
        public List<string> FailedFloors { get; set; } = new List<string>();
        public int TotalSleevesPlaced { get; set; }
        public int TotalClustersFormed { get; set; }
        public TimeSpan TotalDuration { get; set; }
        
        public MultiFloorResult() { }
        
        public MultiFloorResult(IEnumerable<FloorProcessingResult> results)
        {
            foreach (var result in results)
            {
                if (result.Success)
                {
                    SuccessfulFloors.Add(result.LevelName);
                    TotalSleevesPlaced += result.PlacedCount;
                    TotalClustersFormed += result.ClusterCount;
                }
                else
                {
                    FailedFloors.Add(result.LevelName);
                }
            }
        }
        
        public void Merge(MultiFloorResult other)
        {
            SuccessfulFloors.AddRange(other.SuccessfulFloors);
            FailedFloors.AddRange(other.FailedFloors);
            TotalSleevesPlaced += other.TotalSleevesPlaced;
            TotalClustersFormed += other.TotalClustersFormed;
        }
    }
    
    public class FloorProcessingResult
    {
        public string LevelName { get; set; }
        public bool Success { get; set; }
        public int PlacedCount { get; set; }
        public int ClusterCount { get; set; }
        public string? ErrorMessage { get; set; }
        public string? Message { get; set; }
        public List<int> PlacedElements { get; set; } = new List<int>();
    }
    
    public class ValidationResult
    {
        public bool IsValid { get; set; }
        public List<string> Errors { get; set; } = new List<string>();
    }
    
    public class MultiFloorProgress
    {
        public List<string> CompletedFloors { get; set; } = new List<string>();
        public int TotalFloors { get; set; }
    }
    
    public class Checkpoint
    {
        public DateTime Timestamp { get; set; }
        public List<string> ProcessedFloors { get; set; } = new List<string>();
        public int TotalFloors { get; set; }
    }
    
    public class PlacedSleeveData
    {
        public int ElementId { get; set; }
        public string? ZoneGuid { get; set; }
        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }
    }

    public class FloorRawData
    {
        public string LevelName { get; set; }
        public List<ClashZone> Clashes { get; set; } = new List<ClashZone>();
        // Add other raw data fields as needed for post-processing
    }

    public class FloorPostProcessorResult
    {
        public List<FloorProcessingResult> FloorResults { get; set; } = new List<FloorProcessingResult>();
    }
}
