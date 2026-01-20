namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces
{
    /// <summary>
    /// Interface for normalizing file names and category names.
    /// Extracted from UniversalSleevePlacementCommand to adhere to SRP.
    /// </summary>
    public interface IFileNameNormalizer
    {
        /// <summary>
        /// Normalize a file name by removing parentheses, extensions, and element counts.
        /// </summary>
        string NormalizeFileName(string fileName);
        
        /// <summary>
        /// Normalize a category name to match clash zone file naming conventions.
        /// </summary>
        string NormalizeCategoryName(string categoryName);
    }
}

