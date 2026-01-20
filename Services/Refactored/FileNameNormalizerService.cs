using System.IO;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Refactored
{
    /// <summary>
    /// ✅ SOLID REFACTORED: Service for normalizing file names and category names.
    /// Extracted from UniversalSleevePlacementCommand to adhere to SRP.
    /// </summary>
    public class FileNameNormalizerService : IFileNameNormalizer
    {
        public string NormalizeFileName(string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
                return string.Empty;

            // Extract filename from full path (e.g., "C:\Users\...\ME-00001.rvt" -> "ME-00001.rvt")
            fileName = Path.GetFileName(fileName);

            // Remove file extension (e.g., "ME-00001.rvt" -> "ME-00001")
            fileName = Path.GetFileNameWithoutExtension(fileName);

            // Remove content in parentheses (e.g., "Building (Architectural)" -> "Building")
            var idxParen = fileName.IndexOf('(');
            if (idxParen >= 0)
            {
                fileName = fileName.Substring(0, idxParen).Trim();
            }

            // Remove element count suffixes (e.g., "ME-00001 (118 elements)" -> "ME-00001")
            var idxElements = fileName.IndexOf(" elements");
            if (idxElements >= 0)
            {
                fileName = fileName.Substring(0, idxElements).Trim();
            }

            // Remove any trailing spaces and return
            return fileName.Trim();
        }

        public string NormalizeCategoryName(string categoryName)
        {
            if (string.IsNullOrEmpty(categoryName))
                return "unknown";

            return categoryName
                .Replace(" ", "_")
                .ToLowerInvariant();
        }
    }
}

