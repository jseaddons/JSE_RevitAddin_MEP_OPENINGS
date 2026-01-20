using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.ParameterCapture
{
    /// <summary>
    /// SOLID Principle: Single Responsibility Principle (SRP)
    /// Responsibility: Persist learned parameter keys to/from disk.
    /// 
    /// Features Preserved (28-point compliance):
    /// - FEATURE 28: Project-level learned keys persistence
    /// - Thread-safe operations
    /// - Error handling and graceful degradation
    /// </summary>
    public class FileParameterKeyStore : IParameterKeyStore
    {
        private static readonly object _lock = new object();

        /// <summary>
        /// Get the file path for learned keys storage.
        /// FEATURE 28: Project-level persistence in filters directory.
        /// </summary>
        private static string GetLearnedKeysFilePath()
        {
            var dir = ProjectPathService.GetFiltersDirectory(null);
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
            return Path.Combine(dir, "learned_parameter_keys.xml");
        }

        /// <summary>
        /// FEATURE 28: Load learned keys from disk with error handling.
        /// </summary>
        public HashSet<string> LoadLearnedKeys()
        {
            lock (_lock)
            {
                try
                {
                    var file = GetLearnedKeysFilePath();
                    if (!File.Exists(file))
                        return new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                    var doc = new System.Xml.XmlDocument();
                    doc.Load(file);

                    var nodes = doc.SelectNodes("/LearnedParameterKeys/Key");
                    var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                    if (nodes != null)
                    {
                        foreach (System.Xml.XmlNode n in nodes)
                        {
                            var value = n.InnerText?.Trim();
                            if (!string.IsNullOrEmpty(value))
                                keys.Add(value);
                        }
                    }

                    return keys;
                }
                catch (Exception ex)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Error($"[FileParameterKeyStore] Error loading learned keys: {ex.Message}");
                    }
                    return new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                }
            }
        }

        /// <summary>
        /// FEATURE 28: Save learned keys to disk with error handling.
        /// </summary>
        public void SaveLearnedKeys(HashSet<string> keys)
        {
            if (keys == null || keys.Count == 0) return;

            lock (_lock)
            {
                try
                {
                    var file = GetLearnedKeysFilePath();
                    var xml = new System.Xml.XmlDocument();
                    var root = xml.CreateElement("LearnedParameterKeys");
                    xml.AppendChild(root);

                    foreach (var key in keys.OrderBy(s => s, StringComparer.OrdinalIgnoreCase))
                    {
                        var element = xml.CreateElement("Key");
                        element.InnerText = key;
                        root.AppendChild(element);
                    }

                    xml.Save(file);
                }
                catch (Exception ex)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Error($"[FileParameterKeyStore] Error saving learned keys: {ex.Message}");
                    }
                }
            }
        }

        /// <summary>
        /// FEATURE 28: Add a single learned key with deduplication.
        /// </summary>
        public void AddLearnedKey(string key)
        {
            if (string.IsNullOrWhiteSpace(key)) return;

            lock (_lock)
            {
                try
                {
                    var keys = LoadLearnedKeys();
                    if (!keys.Contains(key))
                    {
                        keys.Add(key);
                        SaveLearnedKeys(keys);
                    }
                }
                catch (Exception ex)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Error($"[FileParameterKeyStore] Error adding learned key '{key}': {ex.Message}");
                    }
                }
            }
        }
    }
}
