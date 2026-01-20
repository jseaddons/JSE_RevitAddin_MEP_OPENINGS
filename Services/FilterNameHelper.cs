using System;
using System.IO;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Provides shared helpers for normalizing filter names so Global XML and
    /// category snapshots stay under a single canonical branch.
    /// </summary>
    internal static class FilterNameHelper
    {
        private static readonly string[] KnownCategorySuffixes =
        {
            "_ducts",
            "_pipes",
            "_cable_trays",
            "_cabletrays",
            "_duct_accessories",
            "_ductaccessories",
            "_ductaccessory",
            "_pipe_accessories",
            "_pipeaccessories",
            "_pipeaccessory"
        };

        /// <summary>
        /// Normalizes the base filter name by removing file extensions and
        /// category suffixes (e.g. "Plumbing_pipes" → "Plumbing").
        /// </summary>
        public static string NormalizeBaseName(string primary, string fallback = null, string category = null)
        {
            var source = !string.IsNullOrWhiteSpace(primary) ? primary : fallback;
            if (string.IsNullOrWhiteSpace(source))
                return "Unknown";

            source = source.Trim();
            source = Path.GetFileNameWithoutExtension(source);

            if (string.IsNullOrWhiteSpace(source))
                return "Unknown";

            source = RemoveKnownSuffixes(source);

            if (!string.IsNullOrWhiteSpace(category))
            {
                var normalizedCategory = category.Trim().ToLowerInvariant().Replace(" ", "_");
                if (!string.IsNullOrEmpty(normalizedCategory) &&
                    source.EndsWith("_" + normalizedCategory, StringComparison.OrdinalIgnoreCase))
                {
                    source = source.Substring(0, source.Length - normalizedCategory.Length - 1);
                }
            }

            return string.IsNullOrWhiteSpace(source) ? "Unknown" : source;
        }

        private static string RemoveKnownSuffixes(string value)
        {
            foreach (var suffix in KnownCategorySuffixes)
            {
                if (value.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
                {
                    value = value.Substring(0, value.Length - suffix.Length);
                    break;
                }
            }

            return value;
        }
    }
}

