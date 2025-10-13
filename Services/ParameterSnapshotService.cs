using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.IO;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// OOP service responsible for building parameter whitelists and capturing element parameter snapshots.
    /// Keeps Refresh integration minimal and isolated.
    /// </summary>
    public class ParameterSnapshotService
    {
        private readonly ISet<string> _commonMepKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Size","Diameter","Nominal Diameter","Width","Height",
            "Reference Level","Level","Schedule Level","Reference Level Elevation",
            "System Type","System Classification","Service Type","System Abbreviation"
        };

        private readonly ISet<string> _commonHostKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Fire Rating","Room Name","Room Number"
        };

        /// <summary>
        /// Build a whitelist for the current run by combining curated keys and previously learned keys from storage.
        /// </summary>
        public HashSet<string> BuildWhitelist(ClashZoneStorage storage, IEnumerable<(Element mep, Element host)> sample)
        {
            var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var k in _commonMepKeys) keys.Add(k);
            foreach (var k in _commonHostKeys) keys.Add(k);

            if (storage?.ParameterKeyWhitelist != null)
            {
                foreach (var k in storage.ParameterKeyWhitelist) keys.Add(k);
            }
            if (storage?.LearnedParameterKeys != null)
            {
                foreach (var k in storage.LearnedParameterKeys) keys.Add(k);
            }

            // Merge disk-learned keys (project-level)
            foreach (var k in LoadLearnedKeysFromDisk()) keys.Add(k);

            return keys;
        }

        /// <summary>
        /// Capture whitelisted parameter values for an element (tries instance, then type if missing).
        /// </summary>
        public List<SerializableKeyValue> CaptureParams(Element element, HashSet<string> whitelist)
        {
            var result = new List<SerializableKeyValue>();
            if (element == null || whitelist == null || whitelist.Count == 0) return result;

            foreach (var key in whitelist)
            {
                var p = LookupParam(element, key);
                if (p == null) continue;

                var value = ConvertParameterToString(p);
                if (string.IsNullOrWhiteSpace(value)) continue;

                result.Add(new SerializableKeyValue { Key = key, Value = value });
            }

            return result;
        }

        /// <summary>
        /// Convert a parameter value to a robust invariant string.
        /// </summary>
        private string ConvertParameterToString(Parameter p)
        {
            if (p == null) return string.Empty;

            string value = p.AsString();
            if (!string.IsNullOrEmpty(value)) return value;

            value = p.AsValueString();
            if (!string.IsNullOrEmpty(value)) return value;

            switch (p.StorageType)
            {
                case StorageType.Integer:
                    return p.AsInteger().ToString(CultureInfo.InvariantCulture);
                case StorageType.Double:
                    return p.AsDouble().ToString(CultureInfo.InvariantCulture);
                case StorageType.ElementId:
                    return p.AsElementId()?.IntegerValue.ToString(CultureInfo.InvariantCulture) ?? string.Empty;
                default:
                    return string.Empty;
            }
        }

        private Parameter LookupParam(Element e, string key)
        {
            var p = e.LookupParameter(key);
            if (p != null) return p;

            if (e is FamilyInstance fi)
            {
                var sp = fi.Symbol?.LookupParameter(key);
                if (sp != null) return sp;
            }

            return null;
        }

        /// <summary>
        /// Utility to produce a doc key (host vs linked distinction).
        /// </summary>
        public string GetDocKey(Element e)
        {
            try { return e?.Document?.PathName ?? "Unknown"; }
            catch { return "Unknown"; }
        }

        // === Learned Keys (Project-level Persistence) ===
        private static string GetLearnedKeysFilePath()
        {
            var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings", "Projects", "Default", "Filters");
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
            return Path.Combine(dir, "learned_parameter_keys.xml");
        }

        public IEnumerable<string> LoadLearnedKeysFromDisk()
        {
            try
            {
                var file = GetLearnedKeysFilePath();
                if (!File.Exists(file)) return Enumerable.Empty<string>();

                var doc = new System.Xml.XmlDocument();
                doc.Load(file);
                var nodes = doc.SelectNodes("/LearnedParameterKeys/Key");
                var list = new List<string>();
                if (nodes != null)
                {
                    foreach (System.Xml.XmlNode n in nodes)
                    {
                        var v = n.InnerText?.Trim();
                        if (!string.IsNullOrEmpty(v)) list.Add(v);
                    }
                }
                return list;
            }
            catch { return Enumerable.Empty<string>(); }
        }

        public static void AddLearnedKey(string key)
        {
            if (string.IsNullOrWhiteSpace(key)) return;
            try
            {
                var file = GetLearnedKeysFilePath();
                var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                if (File.Exists(file))
                {
                    var doc = new System.Xml.XmlDocument();
                    doc.Load(file);
                    var nodes = doc.SelectNodes("/LearnedParameterKeys/Key");
                    if (nodes != null)
                    {
                        foreach (System.Xml.XmlNode n in nodes)
                        {
                            var v = n.InnerText?.Trim();
                            if (!string.IsNullOrEmpty(v)) keys.Add(v);
                        }
                    }
                }

                if (!keys.Contains(key)) keys.Add(key);

                // Write out
                var xml = new System.Xml.XmlDocument();
                var root = xml.CreateElement("LearnedParameterKeys");
                xml.AppendChild(root);
                foreach (var k in keys.OrderBy(s => s, StringComparer.OrdinalIgnoreCase))
                {
                    var e = xml.CreateElement("Key");
                    e.InnerText = k;
                    root.AppendChild(e);
                }
                xml.Save(file);
            }
            catch { /* non-fatal */ }
        }
    }
}


