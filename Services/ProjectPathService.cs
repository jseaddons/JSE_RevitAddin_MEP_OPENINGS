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
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var root = Path.Combine(appData, "JSE_MEP_Openings", "Projects");

            // Prefer actual file name from PathName (handles detached/renamed titles)
            string pathName = doc?.PathName;
            string nameFromPath = null;
            if (!string.IsNullOrWhiteSpace(pathName))
            {
                try
                {
                    nameFromPath = Path.GetFileNameWithoutExtension(pathName);
                }
                catch { }
            }

            var sourceName = !string.IsNullOrWhiteSpace(nameFromPath) ? nameFromPath : (doc?.Title ?? "Default");
            var projectName = Sanitize(sourceName);

            var projectRoot = Path.Combine(root, projectName);
            // Lightweight trace for debugging where files go
            try { System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", $"[{DateTime.Now:HH:mm:ss}] PROJECT_ROOT={projectRoot} (source='{sourceName}')\n"); } catch { }
            return projectRoot;
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
