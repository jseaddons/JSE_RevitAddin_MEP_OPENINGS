using System;
using System.Collections.Generic;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Parameters.Configuration
{
    public class MarkPrefixSettings
    {
        public string ProjectPrefix { get; set; } = "SLEEVE_";
        public string DuctPrefix { get; set; } = "DCT";
        public string PipePrefix { get; set; } = "PLU";
        public string CableTrayPrefix { get; set; } = "ELE";
        public string DamperPrefix { get; set; } = "DAM";
        
        public bool RemarkAll { get; set; } = false;
        public string NumberFormat { get; set; } = "000";

        // Legacy properties for serialization compatibility
        public string RemarkDuctPrefix { get { return DuctPrefix; } set { DuctPrefix = value; } }
        public string RemarkPipePrefix { get { return PipePrefix; } set { PipePrefix = value; } }
        public string RemarkCableTrayPrefix { get { return CableTrayPrefix; } set { CableTrayPrefix = value; } }
        public string RemarkDamperPrefix { get { return DamperPrefix; } set { DamperPrefix = value; } }

        public Dictionary<string, string> DuctSystemTypeOverrides { get; set; } = new Dictionary<string, string>();
        public Dictionary<string, string> PipeSystemTypeOverrides { get; set; } = new Dictionary<string, string>();
        
        // Dummy method to satisfy DisciplinePrefixStrategy
        public string GetPrefixForElement(string category, string systemType, string serviceType)
        {
            return "OPN"; // Default fallback
        }

        // Dummy method for MarkParameterCommand (if any usages remain)
        public string GetDisciplinePrefix(string category) => "MEP";
        public bool GetRemarkFlag(string category) => false;
        public bool ActiveViewOnly { get; set; } = false;
    }
}
