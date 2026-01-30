using System;

namespace JSE_RevitAddin_MEP_OPENINGS.Models
{
    /// <summary>
    /// Represents a unique combination of a linked file and a host file that has been processed.
    /// Used for tracking processed files in the global index.
    /// </summary>
    public class ProcessedFileCombo
    {
        public string? LinkedFile { get; set; }
        public string? HostFile { get; set; }

        public override string ToString()
        {
            return $"Link: {LinkedFile ?? "None"}, Host: {HostFile ?? "None"}";
        }

        public string GetNormalizedKey()
        {
            return $"{Normalize(LinkedFile)}_{Normalize(HostFile)}";
        }

        private static string Normalize(string? path)
        {
            if (string.IsNullOrWhiteSpace(path)) return "none";
            
            var name = path;
            var idxParen = name.IndexOf('(');
            if (idxParen >= 0) name = name.Substring(0, idxParen);
            
            try 
            {
                name = System.IO.Path.GetFileNameWithoutExtension(name);
            }
            catch 
            {
                // In case of invalid path characters
            }

            name = name.ToLowerInvariant()
                .Replace("_detached", "")
                .Replace('_', ' ')
                .Replace('-', ' ');
            
            return System.Text.RegularExpressions.Regex.Replace(name, @"\s+", " ").Trim();
        }
    }
}
