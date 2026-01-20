using System.Collections.Generic;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.ParameterCapture
{
    /// <summary>
    /// SOLID Principle: Single Responsibility Principle (SRP)
    /// Handles persistence of learned parameter keys.
    /// Separates storage concerns from capture mechanics and policy.
    /// </summary>
    public interface IParameterKeyStore
    {
        /// <summary>
        /// Load learned parameter keys from persistent storage.
        /// </summary>
        HashSet<string> LoadLearnedKeys();

        /// <summary>
        /// Save learned parameter keys to persistent storage.
        /// </summary>
        void SaveLearnedKeys(HashSet<string> keys);

        /// <summary>
        /// Add a single learned key to persistent storage.
        /// </summary>
        void AddLearnedKey(string key);
    }
}
