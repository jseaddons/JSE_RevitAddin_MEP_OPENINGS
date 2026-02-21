using System;
using System.IO;
using System.Text.RegularExpressions;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    public static class ProjectPathService
    {
        private static string Sanitize(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "Default";
            // Remove invalid path chars and trim
            var invalid = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());
            var pattern = "[" + Regex.Escape(invalid) + "]";
            var cleaned = Regex.Replace(name, pattern, "_");
            cleaned = cleaned.Trim();
            if (cleaned.Length == 0) cleaned = "Default";
            return cleaned;
        }

        public static string GetProjectRoot(Document doc)
        {
            // USER REQUEST: Always use AppData path for consistency between Local and BIM 360 models
            // Path: %APPDATA%\JSE_MEP_Openings\Projects\[ProjectName]
            
            try
            {
                var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                var root = Path.Combine(appData, "JSE_MEP_Openings", "Projects");
                
                // Generate safe project name
                var rawName = doc?.Title;
                if (string.IsNullOrEmpty(rawName)) rawName = "Default_Project";
                
                // Remove file extension if present (e.g. .rvt)
                if (rawName.EndsWith(".rvt", StringComparison.OrdinalIgnoreCase))
                    rawName = Path.GetFileNameWithoutExtension(rawName);
                    
                var projectName = Sanitize(rawName);
                var projectPath = Path.Combine(root, projectName);
                
                // 🔍 DIAGNOSTIC: Log the resolved path
                try {
                     System.Diagnostics.Debug.WriteLine($"ProjectPathService: Resolved '{doc?.Title}' to '{projectPath}'");
                } catch {}

                if (!Directory.Exists(projectPath))
                {
                    Directory.CreateDirectory(projectPath);
                }
                
                return projectPath;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ProjectPathService: Error resolving project path: {ex.Message}");
                // Absolute fallback
                return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings", "Default");
            }
        }

        public static string GetFiltersDirectory(Document doc)
        {
            return Path.Combine(GetProjectRoot(doc), "Filters");
        }

        public static void EnsureFiltersDirectory(Document doc)
        {
            var dir = GetFiltersDirectory(doc);
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
        }
    }
}


