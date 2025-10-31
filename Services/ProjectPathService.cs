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
            var projectName = Sanitize(doc?.Title ?? "Default");
            return Path.Combine(root, projectName);
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


