using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Serialization;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
	[XmlRoot("CategoryGlobalIndex")] 
	public class CategoryGlobalIndex
	{
		[XmlAttribute("Category")] public string Category { get; set; } = string.Empty;
		[XmlElement("Entry")] public List<CategoryGlobalIndexEntry> Entries { get; set; } = new List<CategoryGlobalIndexEntry>();
	}

	public class CategoryGlobalIndexEntry
	{
		[XmlAttribute("Id")] public string Id { get; set; } = string.Empty; // Guid as string
		[XmlAttribute("IsResolved")] public bool IsResolved { get; set; }
		[XmlAttribute("IsClusterResolved")] public bool IsClusterResolved { get; set; }
		[XmlAttribute("SleeveInstanceId")] public int SleeveInstanceId { get; set; }
		[XmlAttribute("ClusterSleeveInstanceId")] public int ClusterSleeveInstanceId { get; set; }
	}

	public static class GlobalIndexService
	{
		private static readonly XmlSerializer _serializer = new XmlSerializer(typeof(CategoryGlobalIndex));

		public static string GetCategoryIndexPath(Document doc, string categoryName)
		{
			var dir = ProjectPathService.GetFiltersDirectory(doc);
			return Path.Combine(dir, $"{Sanitize(categoryName)}_global.xml");
		}

		public static CategoryGlobalIndex LoadOrCreate(Document doc, string categoryName)
		{
			string path = GetCategoryIndexPath(doc, categoryName);
			try
			{
				if (File.Exists(path))
				{
					using (var reader = new StreamReader(path))
					{
						return (CategoryGlobalIndex)_serializer.Deserialize(reader);
					}
				}
			}
			catch { }

			return new CategoryGlobalIndex { Category = categoryName, Entries = new List<CategoryGlobalIndexEntry>() };
		}

		public static void EnsureEntries(Document doc, string categoryName, IEnumerable<Guid> clashZoneIds)
		{
			var index = LoadOrCreate(doc, categoryName);
			bool changed = false;
			var set = new HashSet<string>(index.Entries.Select(e => e.Id), StringComparer.OrdinalIgnoreCase);
			foreach (var id in clashZoneIds)
			{
				string s = id.ToString();
				if (!set.Contains(s))
				{
					index.Entries.Add(new CategoryGlobalIndexEntry { Id = s, IsResolved = false, IsClusterResolved = false, SleeveInstanceId = 0, ClusterSleeveInstanceId = 0 });
					changed = true;
				}
			}
			if (changed)
			{
				Save(doc, index);
			}
		}

		public static void UpsertFlags(Document doc, string categoryName, IEnumerable<(Guid Id, bool IsResolved, bool IsClusterResolved)> updates)
		{
			var index = LoadOrCreate(doc, categoryName);
			var map = index.Entries.ToDictionary(e => e.Id, StringComparer.OrdinalIgnoreCase);
			bool changed = false;
			foreach (var up in updates)
			{
				string key = up.Id.ToString();
				if (!map.TryGetValue(key, out var entry))
				{
					entry = new CategoryGlobalIndexEntry { Id = key };
					index.Entries.Add(entry);
					map[key] = entry;
				}
				if (entry.IsResolved != up.IsResolved || entry.IsClusterResolved != up.IsClusterResolved)
				{
					entry.IsResolved = up.IsResolved;
					entry.IsClusterResolved = up.IsClusterResolved;
					changed = true;
				}
			}
			if (changed)
			{
				Save(doc, index);
			}
		}

		public static void UpsertFlagsWithIds(Document doc, string categoryName, IEnumerable<(Guid Id, bool IsResolved, bool IsClusterResolved, int SleeveInstanceId, int ClusterSleeveInstanceId)> updates)
		{
			var index = LoadOrCreate(doc, categoryName);
			var map = index.Entries.ToDictionary(e => e.Id, StringComparer.OrdinalIgnoreCase);
			bool changed = false;
			foreach (var up in updates)
			{
				string key = up.Id.ToString();
				if (!map.TryGetValue(key, out var entry))
				{
					entry = new CategoryGlobalIndexEntry { Id = key };
					index.Entries.Add(entry);
					map[key] = entry;
				}
				if (entry.IsResolved != up.IsResolved || entry.IsClusterResolved != up.IsClusterResolved || entry.SleeveInstanceId != up.SleeveInstanceId || entry.ClusterSleeveInstanceId != up.ClusterSleeveInstanceId)
				{
					entry.IsResolved = up.IsResolved;
					entry.IsClusterResolved = up.IsClusterResolved;
					entry.SleeveInstanceId = up.SleeveInstanceId;
					entry.ClusterSleeveInstanceId = up.ClusterSleeveInstanceId;
					changed = true;
				}
			}
			if (changed)
			{
				Save(doc, index);
			}
		}

		public static (HashSet<Guid> resolved, HashSet<Guid> clusterResolved) ValidateAndFixFlags(Document doc, string categoryName)
		{
			var index = LoadOrCreate(doc, categoryName);
			bool changed = false;
			var resolved = new HashSet<Guid>();
			var clusterResolved = new HashSet<Guid>();
			foreach (var e in index.Entries)
			{
				Guid id;
				if (!Guid.TryParse(e.Id, out id)) continue;
				// Individual check
				if (e.IsResolved)
				{
					var el = e.SleeveInstanceId > 0 ? doc?.GetElement(new ElementId(e.SleeveInstanceId)) : null;
					if (e.SleeveInstanceId <= 0 || el == null)
					{
						e.IsResolved = false;
						e.SleeveInstanceId = 0;
						changed = true;
					}
					else
					{
						resolved.Add(id);
					}
				}
				// Cluster check
				if (e.IsClusterResolved)
				{
					var el = e.ClusterSleeveInstanceId > 0 ? doc?.GetElement(new ElementId(e.ClusterSleeveInstanceId)) : null;
					if (e.ClusterSleeveInstanceId <= 0 || el == null)
					{
						e.IsClusterResolved = false;
						e.ClusterSleeveInstanceId = 0;
						changed = true;
					}
					else
					{
						clusterResolved.Add(id);
					}
				}
			}
			if (changed)
			{
				Save(doc, index);
			}
			return (resolved, clusterResolved);
		}

		private static void Save(Document doc, CategoryGlobalIndex index)
		{
			try
			{
				ProjectPathService.EnsureFiltersDirectory(doc);
				string path = GetCategoryIndexPath(doc, index.Category);
				using (var writer = new StreamWriter(path))
				{
					_serializer.Serialize(writer, index);
				}
			}
			catch (Exception ex)
			{
				SafeFileLogger.SafeAppendText("Refresh_GlobalIndex.log", $"Failed to save global index for '{index.Category}': {ex.Message}");
			}
		}

		private static string Sanitize(string name)
		{
			foreach (var c in Path.GetInvalidFileNameChars())
			{
				name = name.Replace(c, '_');
			}
			return name;
		}
	}
}
