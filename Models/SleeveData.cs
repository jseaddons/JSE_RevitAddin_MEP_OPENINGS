using System;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Models
{
    /// <summary>
    /// Model for storing actual sleeve data for clustering
    /// </summary>
    public class SleeveData
    {
        public int SleeveInstanceId { get; set; }
        public XYZ Corner1 { get; set; }
        public XYZ Corner2 { get; set; }
        public XYZ Corner3 { get; set; }
        public XYZ Corner4 { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public double Depth { get; set; }
        public string HostType { get; set; }
        public string Orientation { get; set; }
        public string Category { get; set; }
        public DateTime CreatedAt { get; set; }
        
        public SleeveData()
        {
            CreatedAt = DateTime.Now;
        }
    }
}





