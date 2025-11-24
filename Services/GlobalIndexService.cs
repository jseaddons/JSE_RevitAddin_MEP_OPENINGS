using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Serialization;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
	/// <summary>
	/// ✅ HIERARCHICAL STRUCTURE: Represents a file combination (LinkedFile + HostFile) within a filter
	/// Contains clash zones (Entries) for this specific file combo in this filter
	/// </summary>
	public class FileComboGroup
	{
		[XmlAttribute("LinkedFile")] public string LinkedFile { get; set; } = string.Empty;
		[XmlAttribute("HostFile")] public string HostFile { get; set; } = string.Empty;
		[XmlAttribute("ProcessedAt")] public DateTime ProcessedAt { get; set; } = DateTime.Now;
		[XmlAttribute("IsProcessed")] public bool IsProcessed { get; set; } = false; // false = new combo (needs detection), true = existing combo (skip detection)
		
		/// <summary>
		/// Clash zones for this file combo in this filter
		/// </summary>
		[XmlElement("Entry")] public List<CategoryGlobalIndexEntry> Entries { get; set; } = new List<CategoryGlobalIndexEntry>();
		
		/// <summary>
		/// Creates a normalized key for comparison (case-insensitive, removes path info)
		/// ✅ FIX: Handles old format ": X : location Shared" pattern
		/// </summary>
		public string GetNormalizedKey()
		{
			Func<string, string> norm = s =>
			{
				if (string.IsNullOrWhiteSpace(s)) return string.Empty;
				var trimmed = s;
				
				// ✅ FIX: Remove old format ": X : location Shared" pattern (e.g., ": 12 : location Shared")
				// This handles legacy file combo names from Global XML
				var locationMatch = System.Text.RegularExpressions.Regex.Match(trimmed, @":\s*\d+\s*:\s*location\s+Shared", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
				if (locationMatch.Success)
				{
					trimmed = trimmed.Substring(0, locationMatch.Index).Trim();
				}
				
				var idxParen = trimmed.IndexOf('(');
				if (idxParen >= 0) trimmed = trimmed.Substring(0, idxParen);
				trimmed = Path.GetFileNameWithoutExtension(trimmed);
				trimmed = trimmed.ToLowerInvariant().Replace("_detached", "");
				trimmed = trimmed.Replace('_', ' ').Replace('-', ' ');
				trimmed = System.Text.RegularExpressions.Regex.Replace(trimmed, "\\s+", " ");
				return trimmed.Trim();
			};
			
			return $"{norm(LinkedFile)}|{norm(HostFile)}";
		}
	}

	/// <summary>
	/// ✅ HIERARCHICAL STRUCTURE: Represents a filter group containing file combos and their clash zones
	/// Tree structure: Filter → FileCombo → ClashZones
	/// </summary>
	public class FilterGroup
	{
		[XmlAttribute("Name")] public string Name { get; set; } = string.Empty;
		
		/// <summary>
		/// File combinations processed for this filter
		/// </summary>
		[XmlElement("FileCombo")] public List<FileComboGroup> FileCombos { get; set; } = new List<FileComboGroup>();
	}

	/// <summary>
	/// ✅ OPTIMIZATION: Represents a processed file combination (LinkedFile + HostFile)
	/// Used to track which file combos have been processed per category to avoid redundant intersection detection
	/// ⚠️ DEPRECATED: Replaced by hierarchical structure (FilterGroup → FileComboGroup)
	/// Kept for backward compatibility during migration
	/// </summary>
	public class ProcessedFileCombo
	{
		[XmlAttribute("LinkedFile")] public string LinkedFile { get; set; } = string.Empty;
		[XmlAttribute("HostFile")] public string HostFile { get; set; } = string.Empty;
		[XmlAttribute("ProcessedAt")] public DateTime ProcessedAt { get; set; } = DateTime.Now;
		[XmlAttribute("IsProcessed")] public bool IsProcessed { get; set; } = false; // ✅ NEW: false = new combo (needs detection), true = existing combo (skip detection)
		
		/// <summary>
		/// Creates a normalized key for comparison (case-insensitive, removes path info)
		/// ✅ FIX: Handles old format ": X : location Shared" pattern
		/// </summary>
		public string GetNormalizedKey()
		{
			Func<string, string> norm = s =>
			{
				if (string.IsNullOrWhiteSpace(s)) return string.Empty;
				var trimmed = s;
				
				// ✅ FIX: Remove old format ": X : location Shared" pattern (e.g., ": 12 : location Shared")
				// This handles legacy file combo names from Global XML
				var locationMatch = System.Text.RegularExpressions.Regex.Match(trimmed, @":\s*\d+\s*:\s*location\s+Shared", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
				if (locationMatch.Success)
				{
					trimmed = trimmed.Substring(0, locationMatch.Index).Trim();
				}
				
				var idxParen = trimmed.IndexOf('(');
				if (idxParen >= 0) trimmed = trimmed.Substring(0, idxParen);
				trimmed = Path.GetFileNameWithoutExtension(trimmed);
				trimmed = trimmed.ToLowerInvariant().Replace("_detached", "");
				trimmed = trimmed.Replace('_', ' ').Replace('-', ' ');
				trimmed = System.Text.RegularExpressions.Regex.Replace(trimmed, "\\s+", " ");
				return trimmed.Trim();
			};
			
			return $"{norm(LinkedFile)}|{norm(HostFile)}";
		}
	}

	[XmlRoot("CategoryGlobalIndex")] 
	public class CategoryGlobalIndex
	{
		[XmlAttribute("Category")] public string Category { get; set; } = string.Empty;
		
		/// <summary>
		/// ✅ HIERARCHICAL STRUCTURE: Tree organization - Filter → FileCombo → ClashZones
		/// </summary>
		[XmlElement("Filter")] public List<FilterGroup> Filters { get; set; } = new List<FilterGroup>();
		
		/// <summary>
		/// ⚠️ DEPRECATED: Flat structure - kept for backward compatibility during migration
		/// Will be migrated to hierarchical structure (Filters) on first load
		/// </summary>
		[XmlElement("Entry")] public List<CategoryGlobalIndexEntry> Entries { get; set; } = new List<CategoryGlobalIndexEntry>();
		
		/// <summary>
		/// ⚠️ DEPRECATED: Flat structure - kept for backward compatibility during migration
		/// Will be migrated to hierarchical structure (Filters → FileCombos) on first load
		/// </summary>
		[XmlElement("ProcessedFileCombo")] public List<ProcessedFileCombo> ProcessedFileCombos { get; set; } = new List<ProcessedFileCombo>();
	}

	public class CategoryGlobalIndexEntry
	{
		[XmlAttribute("Id")] public string Id { get; set; } = string.Empty; // Guid as string
		[XmlAttribute("IsResolved")] public bool IsResolved { get; set; }
		[XmlAttribute("IsClusterResolved")] public bool IsClusterResolved { get; set; }
		[XmlAttribute("SleeveInstanceId")] public int SleeveInstanceId { get; set; }
		[XmlAttribute("ClusterSleeveInstanceId")] public int ClusterSleeveInstanceId { get; set; }
		[XmlIgnore] public double SleeveWidth { get; set; }
		[XmlIgnore] public double SleeveHeight { get; set; }
		[XmlIgnore] public double SleeveDiameter { get; set; }
		[XmlIgnore] public double SleevePlacementPointX { get; set; }
		[XmlIgnore] public double SleevePlacementPointY { get; set; }
		[XmlIgnore] public double SleevePlacementPointZ { get; set; }
		[XmlIgnore] public double SleevePlacementPointActiveDocumentX { get; set; }
		[XmlIgnore] public double SleevePlacementPointActiveDocumentY { get; set; }
		[XmlIgnore] public double SleevePlacementPointActiveDocumentZ { get; set; }
		[XmlIgnore] public double SleeveBoundingBoxMinX { get; set; }
		[XmlIgnore] public double SleeveBoundingBoxMinY { get; set; }
		[XmlIgnore] public double SleeveBoundingBoxMinZ { get; set; }
		[XmlIgnore] public double SleeveBoundingBoxMaxX { get; set; }
		[XmlIgnore] public double SleeveBoundingBoxMaxY { get; set; }
		[XmlIgnore] public double SleeveBoundingBoxMaxZ { get; set; }
		
		// ✅ CRITICAL: Store MEP+Host+Point for O(1) matching by intersection point (not GUID)
		// This enables cross-filter matching without needing GUID parameter on sleeves
		[XmlAttribute("MepElementId")] public int MepElementId { get; set; }
		[XmlAttribute("StructuralElementId")] public int StructuralElementId { get; set; }
		[XmlAttribute("IntersectionPointX")] public double IntersectionPointX { get; set; }
		[XmlAttribute("IntersectionPointY")] public double IntersectionPointY { get; set; }
		[XmlAttribute("IntersectionPointZ")] public double IntersectionPointZ { get; set; }
		
		/// <summary>
		/// ✅ CRITICAL: Filter name that contains this clash zone's placement data
		/// Used to identify which Filter XML file to load for sleeve placement coordinates, dimensions, host type, etc.
		/// Format: "{FilterName}_{Category}.xml" (e.g., "Pipes_Pipes.xml")
		/// ⚠️ DEPRECATED: In hierarchical structure, FilterName is stored in FilterGroup.Name, not in Entry
		/// Kept for backward compatibility during migration
		/// </summary>
		[XmlAttribute("FilterName")] public string FilterName { get; set; } = string.Empty;
	}

	public static class GlobalIndexService
	{
		private static readonly XmlSerializer _serializer = new XmlSerializer(typeof(CategoryGlobalIndex));

		/// <summary>
		/// ✅ MIGRATION: Converts flat structure (Entries + ProcessedFileCombos) to hierarchical structure (Filters → FileCombos → Entries)
		/// Called automatically on first load of old format XML files
		/// ✅ FIX: Uses actual ProcessedFileCombo values instead of "Migrated" placeholders
		/// </summary>
		private static void MigrateToHierarchicalStructure(CategoryGlobalIndex index)
		{
			// If already migrated (has Filters), skip
			if (index.Filters != null && index.Filters.Count > 0)
				return;
			
			// If no old data to migrate, initialize Filters list
			if (index.Filters == null)
				index.Filters = new List<FilterGroup>();
			
			// If no old entries or processed combos, nothing to migrate
			if ((index.Entries == null || index.Entries.Count == 0) && 
			    (index.ProcessedFileCombos == null || index.ProcessedFileCombos.Count == 0))
				return;
			
			// ✅ STEP 1: Use ProcessedFileCombos to create FileComboGroups with proper LinkedFile/HostFile
			// This ensures we use actual file combo data instead of "Migrated" placeholders
			var fileComboGroups = new Dictionary<string, FileComboGroup>(); // Key: normalized combo key
			
			if (index.ProcessedFileCombos != null && index.ProcessedFileCombos.Count > 0)
			{
				foreach (var combo in index.ProcessedFileCombos)
				{
					if (string.IsNullOrWhiteSpace(combo.LinkedFile) || string.IsNullOrWhiteSpace(combo.HostFile))
						continue; // Skip invalid combos
					
					var normalizedKey = combo.GetNormalizedKey();
					if (!fileComboGroups.ContainsKey(normalizedKey))
					{
						fileComboGroups[normalizedKey] = new FileComboGroup
						{
							LinkedFile = combo.LinkedFile,
							HostFile = combo.HostFile,
							IsProcessed = combo.IsProcessed,
							ProcessedAt = combo.ProcessedAt,
							Entries = new List<CategoryGlobalIndexEntry>()
						};
					}
				}
			}
			
			// ✅ STEP 2: Group entries by FilterName
			var entriesByFilter = (index.Entries ?? new List<CategoryGlobalIndexEntry>())
				.GroupBy(e => e.FilterName ?? string.Empty)
				.ToDictionary(g => g.Key, g => g.ToList());
			
			// ✅ STEP 3: Distribute entries to FileComboGroups ONLY if matching combo exists
			// ✅ CRITICAL FIX: Do NOT create placeholder FileComboGroups with empty LinkedFile/HostFile
			// Entries without matching file combo remain in flat structure only
			foreach (var filterGroup in entriesByFilter)
			{
				var filterName = filterGroup.Key;
				var entries = filterGroup.Value;
				
				// Try to find matching FileComboGroup for this filter
				// Only assign entries to FileComboGroups that exist from ProcessedFileCombos
				FileComboGroup targetCombo = null;
				
				// Try to match entries to existing FileComboGroups (from ProcessedFileCombos)
				if (fileComboGroups.Count > 0)
				{
					// Use the first available FileComboGroup (entries will be assigned to it)
					// This is a best-effort assignment - entries without file combo data are assigned to first combo
					targetCombo = fileComboGroups.Values.FirstOrDefault();
				}
				
				// If no FileComboGroup exists, create one with actual combo data if available from ProcessedFileCombos
				if (targetCombo == null && index.ProcessedFileCombos != null && index.ProcessedFileCombos.Count > 0)
				{
					var firstCombo = index.ProcessedFileCombos.FirstOrDefault(c => 
						!string.IsNullOrWhiteSpace(c.LinkedFile) && !string.IsNullOrWhiteSpace(c.HostFile));
					
					if (firstCombo != null)
					{
						var normalizedKey = firstCombo.GetNormalizedKey();
						targetCombo = new FileComboGroup
						{
							LinkedFile = firstCombo.LinkedFile,
							HostFile = firstCombo.HostFile,
							IsProcessed = firstCombo.IsProcessed,
							ProcessedAt = firstCombo.ProcessedAt,
							Entries = new List<CategoryGlobalIndexEntry>()
						};
						fileComboGroups[normalizedKey] = targetCombo;
					}
				}
				
				// ✅ CRITICAL FIX: Only assign entries to FileComboGroup if we have a valid combo with actual LinkedFile/HostFile
				// DO NOT create placeholder FileComboGroups with empty LinkedFile/HostFile
				// Entries without matching file combo remain in flat structure only
				if (targetCombo != null && !string.IsNullOrWhiteSpace(targetCombo.LinkedFile) && !string.IsNullOrWhiteSpace(targetCombo.HostFile))
				{
					// Assign entries to valid FileComboGroup
					targetCombo.Entries.AddRange(entries);
				}
				// If targetCombo is null or has empty LinkedFile/HostFile, entries remain in flat structure (no assignment)
			}
			
			// ✅ STEP 4: Create FilterGroups and assign FileComboGroups
			var allFilterNames = new HashSet<string>();
			allFilterNames.UnionWith(entriesByFilter.Keys);
			
			// If no filter names but we have file combos, create a default filter
			if (allFilterNames.Count == 0 && fileComboGroups.Count > 0)
			{
				allFilterNames.Add(string.Empty); // Empty filter name for entries without FilterName
			}
			
			foreach (var filterName in allFilterNames)
			{
				var filterGroup = new FilterGroup { Name = filterName, FileCombos = new List<FileComboGroup>() };
				
				// Assign FileComboGroups to this filter
				// For entries with this FilterName, use their associated FileComboGroup
				var entriesForFilter = entriesByFilter.ContainsKey(filterName) ? entriesByFilter[filterName] : new List<CategoryGlobalIndexEntry>();
				
				if (entriesForFilter.Count > 0)
				{
					// ✅ CRITICAL FIX: Do NOT create FileComboGroups during migration unless entries have actual file combo data
					// Entries without file combo data remain in flat structure only
					// FileComboGroups will be created when entries are recreated with proper LinkedFile/HostFile via EnsureEntriesWithClashZoneData
					// 
					// Migration only creates FileComboGroups for ProcessedFileCombos that don't have entries yet
					// (see STEP 5 below)
					//
					// Entries without file combo data stay in flat structure and will be moved to proper FileComboGroups
					// when they're recreated during refresh with SourceDocKey/HostDocKey set
				}
				
				// Also add FileComboGroups that don't have entries yet (if they belong to this filter)
				foreach (var combo in fileComboGroups.Values)
				{
					if (!filterGroup.FileCombos.Contains(combo) && 
					    (combo.Entries == null || combo.Entries.Count == 0))
					{
						filterGroup.FileCombos.Add(combo);
					}
				}
				
				index.Filters.Add(filterGroup);
			}
			
			// ✅ STEP 5: Also create FilterGroups for ProcessedFileCombos that don't have entries
			// This ensures FileComboGroups with proper LinkedFile/HostFile are preserved even without entries
			foreach (var combo in index.ProcessedFileCombos ?? new List<ProcessedFileCombo>())
			{
				if (string.IsNullOrWhiteSpace(combo.LinkedFile) || string.IsNullOrWhiteSpace(combo.HostFile))
					continue;
				
				var normalizedKey = combo.GetNormalizedKey();
				var existingCombo = fileComboGroups.ContainsKey(normalizedKey) ? fileComboGroups[normalizedKey] : null;
				
				// Check if this combo is already in a FilterGroup
				bool alreadyAdded = false;
				foreach (var filter in index.Filters)
				{
					if (filter.FileCombos != null && filter.FileCombos.Any(fc => 
						fc.GetNormalizedKey() == normalizedKey))
					{
						alreadyAdded = true;
						break;
					}
				}
				
				if (!alreadyAdded)
				{
					// Find or create FilterGroup for this combo
					// Use empty filter name if we can't determine it
					var filterGroup = index.Filters.FirstOrDefault(f => string.IsNullOrEmpty(f.Name)) 
						?? new FilterGroup { Name = string.Empty, FileCombos = new List<FileComboGroup>() };
					
					if (!index.Filters.Contains(filterGroup))
						index.Filters.Add(filterGroup);
					
					var fileCombo = existingCombo ?? new FileComboGroup
					{
						LinkedFile = combo.LinkedFile,
						HostFile = combo.HostFile,
						IsProcessed = combo.IsProcessed,
						ProcessedAt = combo.ProcessedAt,
						Entries = new List<CategoryGlobalIndexEntry>()
					};
					
					if (!filterGroup.FileCombos.Contains(fileCombo))
						filterGroup.FileCombos.Add(fileCombo);
				}
			}
		}

		/// <summary>
		/// ✅ HELPER: Gets all entries from hierarchical structure (Filters → FileCombos → Entries) and flat structure (backward compatibility)
		/// </summary>
		public static IEnumerable<CategoryGlobalIndexEntry> GetAllEntries(CategoryGlobalIndex index)
		{
			var allEntries = new List<CategoryGlobalIndexEntry>();
			
			// ✅ HIERARCHICAL STRUCTURE: Collect entries from all filters → file combos
			if (index.Filters != null && index.Filters.Count > 0)
			{
				foreach (var filter in index.Filters)
				{
					if (filter.FileCombos != null)
					{
						foreach (var fileCombo in filter.FileCombos)
						{
							if (fileCombo.Entries != null)
							{
								allEntries.AddRange(fileCombo.Entries);
							}
						}
					}
				}
			}
			
			// ✅ FALLBACK: Also include entries from flat structure (backward compatibility)
			if (index.Entries != null && index.Entries.Count > 0)
			{
				allEntries.AddRange(index.Entries);
			}
			
			return allEntries.Distinct().ToList(); // Remove duplicates by reference
		}

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
						var index = (CategoryGlobalIndex)_serializer.Deserialize(reader);
						
						// ✅ CRITICAL FIX: Ensure hierarchical structure is initialized
						if (index.Filters == null)
							index.Filters = new List<FilterGroup>();
						
						// ✅ CRITICAL FIX: Ensure ProcessedFileCombos is initialized (backward compatibility)
						if (index.ProcessedFileCombos == null)
							index.ProcessedFileCombos = new List<ProcessedFileCombo>();
						
						// ✅ CRITICAL FIX: Ensure Entries is initialized (backward compatibility)
						if (index.Entries == null)
							index.Entries = new List<CategoryGlobalIndexEntry>();
						
						// ✅ CRITICAL FIX: Ensure FilterName is initialized for all entries (backward compatibility)
						if (index.Entries != null)
						{
							foreach (var entry in index.Entries)
							{
								if (entry.FilterName == null)
									entry.FilterName = string.Empty;
							}
						}
						
						// ✅ MIGRATION: Convert flat structure to hierarchical if needed
						MigrateToHierarchicalStructure(index);
						
						return index;
					}
				}
			}
			catch { }

			// ✅ CRITICAL FIX: Initialize hierarchical structure when creating new index
			return new CategoryGlobalIndex 
			{ 
				Category = categoryName, 
				Filters = new List<FilterGroup>(),
				Entries = new List<CategoryGlobalIndexEntry>(), // Keep for backward compatibility
				ProcessedFileCombos = new List<ProcessedFileCombo>() // Keep for backward compatibility
			};
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

		/// <summary>
		/// ✅ CRITICAL: Ensures Global XML entries exist with MEP+Host+Point data for O(1) matching
		/// This enables cross-filter matching without needing GUID parameter on sleeves
		/// ✅ CRITICAL: Stores FilterName to identify which Filter XML file contains placement data
		/// ✅ HIERARCHICAL STRUCTURE: Writes entries to Filter → FileCombo → Entries
		/// </summary>
		/// <param name="doc">Revit document</param>
		/// <param name="categoryName">Category name</param>
		/// <param name="clashZoneData">Clash zone data with GUID, MEP+Host+Point</param>
		/// <param name="filterName">Filter name (required for hierarchical structure)</param>
		/// <param name="linkedFile">Linked file name (optional - if provided, adds to specific FileComboGroup)</param>
		/// <param name="hostFile">Host file name (optional - if provided, adds to specific FileComboGroup)</param>
		public static void EnsureEntriesWithClashZoneData(Document doc, string categoryName, IEnumerable<(Guid Id, int MepElementId, int StructuralElementId, double IntersectionPointX, double IntersectionPointY, double IntersectionPointZ)> clashZoneData, string filterName = null, string linkedFile = null, string hostFile = null)
		{
			var index = LoadOrCreate(doc, categoryName);
			
			// ✅ HIERARCHICAL STRUCTURE: Ensure Filters list is initialized
			if (index.Filters == null)
				index.Filters = new List<FilterGroup>();
			
			// ✅ CRITICAL FIX: Normalize filterName - use empty string if null, but always create FilterGroup
			if (filterName == null)
				filterName = string.Empty;
			
			// ✅ CRITICAL FIX: Always create/get FilterGroup with proper Name (even if empty)
			// This ensures FilterGroup.Name is set immediately when filterName is provided
			FilterGroup filterGroup = index.Filters.FirstOrDefault(f => f.Name == filterName);
			if (filterGroup == null)
			{
				filterGroup = new FilterGroup { Name = filterName, FileCombos = new List<FileComboGroup>() };
				index.Filters.Add(filterGroup);
			}
			
			// ✅ HIERARCHICAL STRUCTURE: Get or create FileComboGroup ONLY if BOTH linkedFile AND hostFile are provided
			// ✅ CRITICAL: Do NOT create placeholder FileComboGroups with empty values
			FileComboGroup fileComboGroup = null;
			if (filterGroup != null && !string.IsNullOrWhiteSpace(linkedFile) && !string.IsNullOrWhiteSpace(hostFile))
			{
				if (filterGroup.FileCombos == null)
					filterGroup.FileCombos = new List<FileComboGroup>();
				
				// Normalize to match existing FileComboGroups
				var combo = new ProcessedFileCombo { LinkedFile = linkedFile, HostFile = hostFile };
				var normalizedKey = combo.GetNormalizedKey();
				
				// ✅ CRITICAL: First check flat ProcessedFileCombos for UI format (most reliable source)
				// This ensures we use the same format that MarkFileCombosAsProcessed uses
				if (index.ProcessedFileCombos != null)
				{
					var existingCombo = index.ProcessedFileCombos.FirstOrDefault(pfc => pfc.GetNormalizedKey() == normalizedKey);
					if (existingCombo != null)
					{
						// Use UI format from existing ProcessedFileCombo
						linkedFile = existingCombo.LinkedFile;
						hostFile = existingCombo.HostFile;
					}
				}
				
				// Find existing FileComboGroup by normalized key in CURRENT FilterGroup first
				fileComboGroup = filterGroup.FileCombos.FirstOrDefault(fc => fc.GetNormalizedKey() == normalizedKey);
				
				// If not found in current FilterGroup, check other FilterGroups for UI format reference
				if (fileComboGroup == null)
				{
					var existingComboInOtherFilter = index.Filters
						.Where(f => f != filterGroup)
						.SelectMany(f => f.FileCombos ?? new List<FileComboGroup>())
						.FirstOrDefault(fc => fc.GetNormalizedKey() == normalizedKey);
					
					if (existingComboInOtherFilter != null)
					{
						// Use UI format from existing FileComboGroup in another filter
						linkedFile = existingComboInOtherFilter.LinkedFile;
						hostFile = existingComboInOtherFilter.HostFile;
					}
				}
				
				// If still not found in current FilterGroup, create new FileComboGroup
				if (fileComboGroup == null)
				{
					// Create new FileComboGroup for this file combo (using UI format if available)
					fileComboGroup = new FileComboGroup
					{
						LinkedFile = linkedFile,
						HostFile = hostFile,
						ProcessedAt = DateTime.Now,
						IsProcessed = true, // Mark as processed since entries are being created for it
						Entries = new List<CategoryGlobalIndexEntry>()
					};
					filterGroup.FileCombos.Add(fileComboGroup);
				}
			}
			
			// ✅ FALLBACK: Keep flat structure for backward compatibility
			if (index.Entries == null)
				index.Entries = new List<CategoryGlobalIndexEntry>();
			
			var set = new HashSet<string>(index.Entries.Select(e => e.Id), StringComparer.OrdinalIgnoreCase);
			bool changed = false;
			
			foreach (var data in clashZoneData)
			{
				string guidString = data.Id.ToString();
				var newEntry = new CategoryGlobalIndexEntry 
					{ 
						Id = guidString, 
						IsResolved = false, 
						IsClusterResolved = false, 
						SleeveInstanceId = 0, 
						ClusterSleeveInstanceId = 0,
						MepElementId = data.MepElementId,
						StructuralElementId = data.StructuralElementId,
						IntersectionPointX = data.IntersectionPointX,
						IntersectionPointY = data.IntersectionPointY,
					IntersectionPointZ = data.IntersectionPointZ,
					FilterName = filterName ?? string.Empty
				};
				
				if (!set.Contains(guidString))
				{
					// ✅ HIERARCHICAL STRUCTURE: Add to FileComboGroup ONLY if LinkedFile/HostFile are provided
					// Do NOT create placeholder FileComboGroups with empty values
					if (fileComboGroup != null)
					{
						if (fileComboGroup.Entries == null)
							fileComboGroup.Entries = new List<CategoryGlobalIndexEntry>();
						
						// Check if entry already exists in FileComboGroup
						if (!fileComboGroup.Entries.Any(e => string.Equals(e.Id, guidString, StringComparison.OrdinalIgnoreCase)))
						{
							fileComboGroup.Entries.Add(newEntry);
							set.Add(guidString); // Mark as added
					changed = true;
						}
				}
				else
				{
						// ✅ CRITICAL: If no FileComboGroup (because LinkedFile/HostFile not provided), entries go to flat structure only
						// Do NOT create placeholder FileComboGroups
						index.Entries.Add(newEntry);
						set.Add(guidString);
						changed = true;
					}
				}
				else
				{
					// ✅ CRITICAL FIX: Update existing entry in BOTH hierarchical and flat structures
					// First check FileComboGroup entries, then flat structure
					CategoryGlobalIndexEntry existingEntry = null;
					
					// Check FileComboGroup entries first
					if (fileComboGroup != null && fileComboGroup.Entries != null)
					{
						existingEntry = fileComboGroup.Entries.FirstOrDefault(e => string.Equals(e.Id, guidString, StringComparison.OrdinalIgnoreCase));
					}
					
					// If not found in FileComboGroup, check flat structure
					if (existingEntry == null)
					{
						existingEntry = index.Entries.FirstOrDefault(e => string.Equals(e.Id, guidString, StringComparison.OrdinalIgnoreCase));
					}
					
					if (existingEntry != null)
					{
						bool entryChanged = false;
						
						// ✅ CRITICAL: Always update MEP+Host+Point data if invalid (0 values)
						if (existingEntry.MepElementId == 0 || existingEntry.StructuralElementId == 0 || 
						    (Math.Abs(existingEntry.IntersectionPointX) < 1e-9 && Math.Abs(existingEntry.IntersectionPointY) < 1e-9 && Math.Abs(existingEntry.IntersectionPointZ) < 1e-9))
						{
							if (data.MepElementId > 0 && data.StructuralElementId > 0)
					{
						existingEntry.MepElementId = data.MepElementId;
						existingEntry.StructuralElementId = data.StructuralElementId;
						existingEntry.IntersectionPointX = data.IntersectionPointX;
						existingEntry.IntersectionPointY = data.IntersectionPointY;
						existingEntry.IntersectionPointZ = data.IntersectionPointZ;
								entryChanged = true;
							}
						}
						
						// ✅ CRITICAL: Always update FilterName if different (set immediately, even if empty)
						if (existingEntry.FilterName != filterName)
						{
							existingEntry.FilterName = filterName ?? string.Empty;
							entryChanged = true;
						}
						
						// ✅ CRITICAL FIX: If entry is in flat structure but LinkedFile/HostFile are provided,
						// MOVE it to FileComboGroup (hierarchical structure)
						if (fileComboGroup != null && index.Entries.Contains(existingEntry))
						{
							// Remove from flat structure
							index.Entries.Remove(existingEntry);
							
							// Add to FileComboGroup
							if (fileComboGroup.Entries == null)
								fileComboGroup.Entries = new List<CategoryGlobalIndexEntry>();
							
							// Check if already in FileComboGroup (shouldn't happen, but safety check)
							if (!fileComboGroup.Entries.Any(e => string.Equals(e.Id, guidString, StringComparison.OrdinalIgnoreCase)))
							{
								fileComboGroup.Entries.Add(existingEntry);
								entryChanged = true;
								
								if (!DeploymentConfiguration.DeploymentMode)
									DebugLogger.Info($"[GLOBAL-INDEX] ✅ MOVED entry {guidString} from flat structure to FileComboGroup (LinkedFile='{linkedFile}', HostFile='{hostFile}')");
							}
						}
						
						if (entryChanged)
						changed = true;
					}
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
			// ✅ HIERARCHICAL STRUCTURE: Use GetAllEntries to search both hierarchical and flat structures
			var allEntries = GetAllEntries(index).ToList();
			var map = allEntries.ToDictionary(e => e.Id, StringComparer.OrdinalIgnoreCase);
			bool changed = false;
			
			foreach (var up in updates)
			{
				string key = up.Id.ToString();
				if (!map.TryGetValue(key, out var entry))
				{
					// Create new entry in flat structure (will be migrated to hierarchical on next load)
					entry = new CategoryGlobalIndexEntry { Id = key };
					index.Entries.Add(entry);
					map[key] = entry;
					changed = true;
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
			// ✅ HIERARCHICAL STRUCTURE: Use GetAllEntries to search both hierarchical and flat structures
			var allEntries = GetAllEntries(index).ToList();
			var map = allEntries.ToDictionary(e => e.Id, StringComparer.OrdinalIgnoreCase);
			bool changed = false;
			
			foreach (var up in updates)
			{
				string key = up.Id.ToString();
				if (!map.TryGetValue(key, out var entry))
				{
					// Create new entry in flat structure (will be migrated to hierarchical on next load)
					entry = new CategoryGlobalIndexEntry { Id = key };
					index.Entries.Add(entry);
					map[key] = entry;
					changed = true;
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

		/// <summary>
		/// ✅ CRITICAL: Upserts flags and sleeve IDs with MEP+Host+Point data for O(1) matching
		/// This enables cross-filter matching without needing GUID parameter on sleeves
	/// ✅ FIX: If GUID doesn't match, tries matching by MEP+Host+Point to avoid duplicate entries
	/// ✅ CRITICAL: Accepts FilterName to identify which Filter XML file contains placement data
		/// </summary>
	public static void UpsertFlagsWithIdsAndClashZoneData(Document doc, string categoryName, IEnumerable<(Guid Id, bool IsResolved, bool IsClusterResolved, int SleeveInstanceId, int ClusterSleeveInstanceId, int MepElementId, int StructuralElementId, double IntersectionPointX, double IntersectionPointY, double IntersectionPointZ)> updates, string filterName = null, string refreshLogName = null)
		{
			// ✅ CRITICAL: UNMISTAKABLE MARKER - This confirms the NEW code is running
			// ✅ ALWAYS LOG TO FILE FIRST (bypasses DebugLogger filtering)
			if (!string.IsNullOrWhiteSpace(refreshLogName))
			{
				try
				{
					SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] ═══════════════════════════════════════════════════════════════════════════════\n");
					SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [GLOBAL-INDEX] ⭐⭐⭐ NEW CODE VERSION v2.0 IS RUNNING ⭐⭐⭐\n");
					SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [GLOBAL-INDEX] ✅ Enhanced logging with entry lookup and update tracking enabled\n");
					SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] ═══════════════════════════════════════════════════════════════════════════════\n");
				}
				catch (Exception fileEx)
				{
					// If file logging fails, try DebugLogger as fallback
					try { DebugLogger.Error($"[GLOBAL-INDEX] Error writing version marker to file: {fileEx.Message}"); } catch { }
				}
			}
			
			DebugLogger.Info("═══════════════════════════════════════════════════════════════════════════════");
			DebugLogger.Info("[GLOBAL-INDEX] ⭐⭐⭐ NEW CODE VERSION v2.0 IS RUNNING ⭐⭐⭐");
			DebugLogger.Info("[GLOBAL-INDEX] ✅ Enhanced logging with entry lookup and update tracking enabled");
			DebugLogger.Info("═══════════════════════════════════════════════════════════════════════════════");
			
			// ✅ BUILD TIMESTAMP: Always log build timestamp (even in deployment mode for troubleshooting)
			try
			{
				var assembly = System.Reflection.Assembly.GetExecutingAssembly();
				string buildTimestamp = assembly.GetName().Version?.ToString() ?? "Unknown";
				DateTime buildTime = DateTime.MinValue;
				string location = assembly.Location;
				if (!string.IsNullOrEmpty(location) && System.IO.File.Exists(location))
				{
					buildTime = System.IO.File.GetLastWriteTime(location);
				}
				else
				{
					// Fallback: Use current time if location unavailable
					buildTime = DateTime.Now;
				}
				DebugLogger.Info($"[GLOBAL-INDEX] ===== BUILD INFO: Version={buildTimestamp}, BuildTime={buildTime:yyyy-MM-dd HH:mm:ss}, Location={location ?? "N/A"}, Method=UpsertFlagsWithIdsAndClashZoneData =====");
			}
			catch (Exception buildEx)
			{
				DebugLogger.Info($"[GLOBAL-INDEX] ===== BUILD INFO: Error getting build info: {buildEx.Message}, Method=UpsertFlagsWithIdsAndClashZoneData =====");
			}
			
			var index = LoadOrCreate(doc, categoryName);

		// ✅ HIERARCHICAL + FLAT: Build a lookup that tracks all entries sharing the same GUID
		var allEntries = GetAllEntries(index).ToList();
		var entriesById = new Dictionary<string, List<CategoryGlobalIndexEntry>>(StringComparer.OrdinalIgnoreCase);

		foreach (var entry in allEntries)
		{
			if (entry == null || string.IsNullOrWhiteSpace(entry.Id))
				continue;

			if (!entriesById.TryGetValue(entry.Id, out var list))
			{
				list = new List<CategoryGlobalIndexEntry>();
				entriesById[entry.Id] = list;
			}
			list.Add(entry);
		}

			bool changed = false;

			foreach (var up in updates)
			{
			if (up.Id == Guid.Empty)
				continue;

				string key = up.Id.ToString();
			List<CategoryGlobalIndexEntry> entryList = null;

			if (!string.IsNullOrWhiteSpace(key))
				entriesById.TryGetValue(key, out entryList);

			// ✅ ENHANCED DEBUG: Log if entry was found (always log, even in deployment mode for troubleshooting)
			if (entryList == null || entryList.Count == 0)
			{
				DebugLogger.Warning($"[GLOBAL-INDEX] ⚠️ Entry {key} NOT FOUND in Global XML (checked {allEntries.Count} total entries) - Update will create new entry in flat structure");
				if (!string.IsNullOrWhiteSpace(refreshLogName))
				{
					try { SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [GLOBAL-INDEX] ⚠️ Entry {key} NOT FOUND in Global XML (checked {allEntries.Count} total entries)\n"); } catch { }
				}
			}
			else
			{
				DebugLogger.Info($"[GLOBAL-INDEX] ✅ Found Entry {key} in Global XML ({entryList.Count} instance(s) - hierarchical + flat)");
				if (!string.IsNullOrWhiteSpace(refreshLogName))
				{
					try { SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [GLOBAL-INDEX] ✅ Found Entry {key} in Global XML ({entryList.Count} instance(s))\n"); } catch { }
				}
			}

			// ✅ STEP 2: If no GUID match, try matching by MEP+Host+Point (handles deterministic GUID collisions)
			if ((entryList == null || entryList.Count == 0) &&
			    up.MepElementId > 0 && up.StructuralElementId > 0 &&
			    Math.Abs(up.IntersectionPointX) > 1e-9 &&
			    Math.Abs(up.IntersectionPointY) > 1e-9 &&
			    Math.Abs(up.IntersectionPointZ) > 1e-9)
			{
				var existingEntry = FindByMepHostAndPoint(doc, categoryName, up.MepElementId, up.StructuralElementId, up.IntersectionPointX, up.IntersectionPointY, up.IntersectionPointZ);
				if (existingEntry != null)
				{
					key = existingEntry.Id;
					if (!entriesById.TryGetValue(key, out entryList) || entryList == null)
					{
						entryList = new List<CategoryGlobalIndexEntry>();
						entriesById[key] = entryList;
					}
					if (!entryList.Contains(existingEntry))
						entryList.Add(existingEntry);

					if (!DeploymentConfiguration.DeploymentMode)
					{
						DebugLogger.Info($"[GLOBAL-INDEX] ✅ MATCHED BY MEP+Host+Point: Entry {existingEntry.Id} reused for incoming GUID {up.Id}");
					}
				}
				}
				
			// ✅ STEP 3: Still no entry? Create one in the flat structure
			if (entryList == null || entryList.Count == 0)
			{
				var newEntry = new CategoryGlobalIndexEntry
				{
					Id = key,
					IsResolved = up.IsResolved,
					IsClusterResolved = up.IsClusterResolved,
					SleeveInstanceId = up.SleeveInstanceId,
					ClusterSleeveInstanceId = up.ClusterSleeveInstanceId,
					MepElementId = up.MepElementId,
					StructuralElementId = up.StructuralElementId,
					IntersectionPointX = up.IntersectionPointX,
					IntersectionPointY = up.IntersectionPointY,
					IntersectionPointZ = up.IntersectionPointZ,
					FilterName = filterName ?? string.Empty
				};

				index.Entries.Add(newEntry);
				entryList = new List<CategoryGlobalIndexEntry> { newEntry };
				entriesById[key] = entryList;
				changed = true;
				continue; // New entry already has correct values; proceed to next update
			}

			// ✅ STEP 4: Update EVERY entry that shares this GUID (hierarchical + flat)
			foreach (var entry in entryList)
			{
				if (entry == null)
					continue;

				bool entryChanged = false;

				// ✅ ENHANCED DEBUG: Log before and after values
				bool wasResolved = entry.IsResolved;
				bool wasClusterResolved = entry.IsClusterResolved;
				int oldSleeveId = entry.SleeveInstanceId;
				int oldClusterId = entry.ClusterSleeveInstanceId;

				// ✅ CRITICAL DEBUG: Always log the update attempt (even in deployment mode for troubleshooting)
				DebugLogger.Info($"[GLOBAL-INDEX] Processing update for Entry {entry.Id}: Current (IsResolved={wasResolved}, IsClusterResolved={wasClusterResolved}, SleeveId={oldSleeveId}, ClusterId={oldClusterId}) → New (IsResolved={up.IsResolved}, IsClusterResolved={up.IsClusterResolved}, SleeveId={up.SleeveInstanceId}, ClusterId={up.ClusterSleeveInstanceId})");
				if (!string.IsNullOrWhiteSpace(refreshLogName))
				{
					try { SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [GLOBAL-INDEX] Processing Entry {entry.Id}: Current (IsResolved={wasResolved}, IsClusterResolved={wasClusterResolved}, SleeveId={oldSleeveId}, ClusterId={oldClusterId}) → New (IsResolved={up.IsResolved}, IsClusterResolved={up.IsClusterResolved}, SleeveId={up.SleeveInstanceId}, ClusterId={up.ClusterSleeveInstanceId})\n"); } catch { }
				}

				if (entry.IsResolved != up.IsResolved || entry.SleeveInstanceId != up.SleeveInstanceId ||
				    entry.IsClusterResolved != up.IsClusterResolved || entry.ClusterSleeveInstanceId != up.ClusterSleeveInstanceId)
				{
					entry.IsResolved = up.IsResolved;
					entry.SleeveInstanceId = up.SleeveInstanceId;
					entry.IsClusterResolved = up.IsClusterResolved;
					entry.ClusterSleeveInstanceId = up.ClusterSleeveInstanceId;
					entryChanged = true;
					
					// ✅ CRITICAL DEBUG: Log the actual update (always log, even in deployment mode)
					DebugLogger.Info($"[GLOBAL-INDEX] ✅ UPDATED Entry {entry.Id}: IsResolved {wasResolved}→{up.IsResolved}, IsClusterResolved {wasClusterResolved}→{up.IsClusterResolved}, SleeveId {oldSleeveId}→{up.SleeveInstanceId}, ClusterId {oldClusterId}→{up.ClusterSleeveInstanceId}");
					if (!string.IsNullOrWhiteSpace(refreshLogName))
					{
						try { SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [GLOBAL-INDEX] ✅ UPDATED Entry {entry.Id}: IsResolved {wasResolved}→{up.IsResolved}, IsClusterResolved {wasClusterResolved}→{up.IsClusterResolved}, SleeveId {oldSleeveId}→{up.SleeveInstanceId}, ClusterId {oldClusterId}→{up.ClusterSleeveInstanceId}\n"); } catch { }
					}
				}
				else
				{
					DebugLogger.Info($"[GLOBAL-INDEX] ⚠️ Entry {entry.Id}: No change needed - IsResolved={entry.IsResolved}, IsClusterResolved={entry.IsClusterResolved} (matches update values)");
					if (!string.IsNullOrWhiteSpace(refreshLogName))
					{
						try { SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [GLOBAL-INDEX] ⚠️ Entry {entry.Id}: No change needed - values match\n"); } catch { }
					}
				}
				
				if (entry.MepElementId != up.MepElementId ||
				    entry.StructuralElementId != up.StructuralElementId ||
				                   Math.Abs(entry.IntersectionPointX - up.IntersectionPointX) > 1e-9 ||
				                   Math.Abs(entry.IntersectionPointY - up.IntersectionPointY) > 1e-9 ||
				    Math.Abs(entry.IntersectionPointZ - up.IntersectionPointZ) > 1e-9)
				{
					entry.MepElementId = up.MepElementId;
					entry.StructuralElementId = up.StructuralElementId;
					entry.IntersectionPointX = up.IntersectionPointX;
					entry.IntersectionPointY = up.IntersectionPointY;
					entry.IntersectionPointZ = up.IntersectionPointZ;
					entryChanged = true;
				}

				if (!string.IsNullOrEmpty(filterName) && entry.FilterName != filterName)
				{
					entry.FilterName = filterName;
					entryChanged = true;
				}

				if (entryChanged)
					changed = true;
				}
			}

			// ✅ CRITICAL DEBUG: Always log whether Save() will be called
			// NOTE: 'changed' means data was modified (entries updated), NOT that flag values are true/false
			int updateCount = updates.Count();
			int clusterUpdates = updates.Count(u => u.IsClusterResolved);
			DebugLogger.Info($"[GLOBAL-INDEX] UpsertFlagsWithIdsAndClashZoneData completed: {updateCount} updates processed ({clusterUpdates} cluster updates), dataModified={changed}, Save() will be called: {changed}");
			if (!string.IsNullOrWhiteSpace(refreshLogName))
			{
				try { SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [GLOBAL-INDEX] Completed: {updateCount} updates processed ({clusterUpdates} cluster updates), dataModified={changed} (NOT flag values), Save() will be called: {changed}\n"); } catch { }
			}
			
			// ✅ CRITICAL FIX: Always save if there are cluster updates, even if 'changed' is false
			// This ensures cluster flags are persisted to XML
			if (changed || clusterUpdates > 0)
			{
				// ✅ CRITICAL: Verify memory state BEFORE Save() to confirm updates are in memory
				var entriesBeforeSave = GetAllEntries(index).ToList();
				int resolvedBeforeSave = entriesBeforeSave.Count(e => e.IsResolved);
				int clusterResolvedBeforeSave = entriesBeforeSave.Count(e => e.IsClusterResolved);
				
				DebugLogger.Info($"[GLOBAL-INDEX] ✅ Calling Save() for category '{categoryName}' - changes will be persisted to XML");
				DebugLogger.Info($"[GLOBAL-INDEX] 📊 MEMORY STATE BEFORE SAVE: IsResolved={resolvedBeforeSave}, IsClusterResolved={clusterResolvedBeforeSave} (out of {entriesBeforeSave.Count} total entries)");
				if (!string.IsNullOrWhiteSpace(refreshLogName))
				{
					try 
					{ 
						SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [GLOBAL-INDEX] ✅ Calling Save() for category '{categoryName}'\n");
						SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [GLOBAL-INDEX] 📊 MEMORY STATE BEFORE SAVE: IsResolved={resolvedBeforeSave}, IsClusterResolved={clusterResolvedBeforeSave} (out of {entriesBeforeSave.Count} total entries)\n");
					} 
					catch { }
				}
				
				try
				{
					Save(doc, index);
					
					// ✅ CRITICAL: Verify file state AFTER Save() to confirm persistence worked
					var filePath = GetCategoryIndexPath(doc, categoryName);
					if (File.Exists(filePath))
					{
						var fileContent = File.ReadAllText(filePath);
						int trueResolvedInFile = (fileContent.Split(new[] { "IsResolved=\"true\"" }, StringSplitOptions.None).Length - 1);
						int falseResolvedInFile = (fileContent.Split(new[] { "IsResolved=\"false\"" }, StringSplitOptions.None).Length - 1);
						int trueClusterInFile = (fileContent.Split(new[] { "IsClusterResolved=\"true\"" }, StringSplitOptions.None).Length - 1);
						int falseClusterInFile = (fileContent.Split(new[] { "IsClusterResolved=\"false\"" }, StringSplitOptions.None).Length - 1);
						
						DebugLogger.Info($"[GLOBAL-INDEX] ✅ Save() completed for category '{categoryName}'");
						DebugLogger.Info($"[GLOBAL-INDEX] 📊 FILE STATE AFTER SAVE: IsResolved true={trueResolvedInFile}, false={falseResolvedInFile} | IsClusterResolved true={trueClusterInFile}, false={falseClusterInFile}");
						DebugLogger.Info($"[GLOBAL-INDEX] 📊 MEMORY STATE AFTER SAVE: IsResolved={resolvedBeforeSave}, IsClusterResolved={clusterResolvedBeforeSave}");
						
						if (trueResolvedInFile != resolvedBeforeSave || trueClusterInFile != clusterResolvedBeforeSave)
						{
							DebugLogger.Error($"[GLOBAL-INDEX] ❌❌❌ PERSISTENCE FAILURE: File flags don't match memory! Memory has {resolvedBeforeSave} IsResolved, {clusterResolvedBeforeSave} IsClusterResolved, but file has {trueResolvedInFile} IsResolved=true, {trueClusterInFile} IsClusterResolved=true");
						}
						
						if (!string.IsNullOrWhiteSpace(refreshLogName))
						{
							try 
							{ 
								SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [GLOBAL-INDEX] ✅ Save() completed for category '{categoryName}'\n");
								SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [GLOBAL-INDEX] 📊 FILE STATE AFTER SAVE: IsResolved true={trueResolvedInFile}, false={falseResolvedInFile} | IsClusterResolved true={trueClusterInFile}, false={falseClusterInFile}\n");
								SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [GLOBAL-INDEX] 📊 MEMORY STATE AFTER SAVE: IsResolved={resolvedBeforeSave}, IsClusterResolved={clusterResolvedBeforeSave}\n");
								if (trueResolvedInFile != resolvedBeforeSave || trueClusterInFile != clusterResolvedBeforeSave)
								{
									SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [GLOBAL-INDEX] ❌❌❌ PERSISTENCE FAILURE: File flags don't match memory!\n");
								}
							} 
							catch { }
						}
					}
					else
					{
						DebugLogger.Error($"[GLOBAL-INDEX] ❌ Save() completed but file does NOT exist at: {filePath}");
						if (!string.IsNullOrWhiteSpace(refreshLogName))
						{
							try { SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [GLOBAL-INDEX] ❌ Save() completed but file does NOT exist at: {filePath}\n"); } catch { }
						}
					}
				}
				catch (Exception saveEx)
				{
					DebugLogger.Error($"[GLOBAL-INDEX] ❌ Save() FAILED for category '{categoryName}': {saveEx.Message}\n{saveEx.StackTrace}");
					if (!string.IsNullOrWhiteSpace(refreshLogName))
					{
						try { SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [GLOBAL-INDEX] ❌ Save() FAILED: {saveEx.Message}\n{saveEx.StackTrace}\n"); } catch { }
					}
					throw; // Re-throw to ensure caller knows Save() failed
				}
			}
			else
			{
				DebugLogger.Warning($"[GLOBAL-INDEX] ⚠️ Save() NOT called for category '{categoryName}' - no changes detected (this may indicate entries weren't found or values matched)");
				if (!string.IsNullOrWhiteSpace(refreshLogName))
				{
					try { SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [GLOBAL-INDEX] ⚠️ Save() NOT called - no changes detected\n"); } catch { }
				}
			}
		}

		/// <summary>
		/// Gets resolved GUID sets from Global XML (no Revit API calls, just reads flags)
		/// Used for filtering during refresh - returns clash zones that are already resolved
		/// Note: Actual flag validation/reset happens:
		/// - 3-Point Validation: When MEP/host elements are deleted (clears Global XML entries)
		/// - ResetResolvedFlagForDeletedSleeves: When sleeves are deleted (resets flags)
		/// </summary>
		public static (HashSet<Guid> resolved, HashSet<Guid> clusterResolved) GetResolvedGuids(Document doc, string categoryName)
		{
			var index = LoadOrCreate(doc, categoryName);
			var resolved = new HashSet<Guid>();
			var clusterResolved = new HashSet<Guid>();
			
			// ✅ HIERARCHICAL STRUCTURE: Use GetAllEntries to read from both hierarchical and flat structures
			var allEntries = GetAllEntries(index);
			
			foreach (var e in allEntries)
			{
				Guid id;
				if (!Guid.TryParse(e.Id, out id)) continue;
				
				// ✅ Cluster resolved: both flags should be true (individual was placed then deleted during clustering)
				if (e.IsClusterResolved)
				{
					clusterResolved.Add(id);
					resolved.Add(id); // Cluster resolved means individual was also resolved (then deleted)
				}
				// ✅ Individual-only resolved: no cluster, but individual flag is true
				else if (e.IsResolved)
					{
						resolved.Add(id);
					}
				// Both flags false = unresolved clash zone (not added to sets)
			}
			
			return (resolved, clusterResolved);
		}

		/// <summary>
		/// Gets all categories that have Global XML files (for checking resolved GUIDs across all categories)
		/// Used when switching filters to ensure all Global XML entries are checked (no expensive Revit API calls)
		/// </summary>
		/// <param name="doc">Revit document</param>
		/// <returns>Set of category names that have Global XML files</returns>
		public static HashSet<string> GetAllCategoriesWithGlobalXml(Document doc)
		{
			var categories = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
			
			try
			{
				var filtersDir = ProjectPathService.GetFiltersDirectory(doc);
				if (!Directory.Exists(filtersDir))
					return categories;
				
				// Find all Global XML files (format: "{Category}_global.xml")
				var globalXmlFiles = Directory.GetFiles(filtersDir, "*_global.xml");
				
				foreach (var xmlFile in globalXmlFiles)
				{
					try
					{
						var fileName = Path.GetFileNameWithoutExtension(xmlFile); // e.g., "Plumbing_global"
						if (fileName.EndsWith("_global", StringComparison.OrdinalIgnoreCase))
						{
							var category = fileName.Substring(0, fileName.Length - "_global".Length);
							if (!string.IsNullOrWhiteSpace(category))
							{
								categories.Add(category);
							}
						}
					}
					catch
					{
						// Skip invalid file names
					}
				}
			}
			catch (Exception ex)
			{
				// Log error but return what we have
				SafeFileLogger.SafeAppendText("Refresh_GlobalIndex.log", 
					$"Error getting categories with Global XML: {ex.Message}");
			}
			
			return categories;
		}

		/// <summary>
		/// Gets resolved GUID sets from Global XML for multiple categories (OOP method to avoid duplicate logic)
		/// Used when filter has multiple categories or new categories are added
		/// Returns a combined HashSet of all resolved GUIDs from all specified categories
		/// </summary>
		/// <param name="doc">Revit document</param>
		/// <param name="categories">Set of category names to check</param>
		/// <returns>Combined HashSet of all resolved GUIDs (both individual and cluster resolved)</returns>
		public static HashSet<Guid> GetResolvedGuidsForCategories(Document doc, HashSet<string> categories)
		{
			var allResolved = new HashSet<Guid>();
			
			if (categories == null || categories.Count == 0)
				return allResolved;
			
			foreach (var category in categories)
			{
				if (string.IsNullOrWhiteSpace(category))
					continue;
				
				try
				{
					var (resolvedSet, clusterResolvedSet) = GetResolvedGuids(doc, category);
					
					// Add both individual and cluster resolved GUIDs
					foreach (var gid in resolvedSet)
						allResolved.Add(gid);
					
					foreach (var gid in clusterResolvedSet)
						allResolved.Add(gid);
				}
				catch (Exception ex)
				{
					// Log error but continue with other categories
					SafeFileLogger.SafeAppendText("Refresh_GlobalIndex.log", 
						$"Error getting resolved GUIDs for category '{category}': {ex.Message}");
				}
			}
			
			return allResolved;
		}

		/// <summary>
		/// ✅ CRITICAL: Finds Global XML entry by MEP+Host+Point (O(1) lookup using Dictionary)
		/// This enables cross-filter matching without needing GUID parameter on sleeves
		/// ✅ HIERARCHICAL STRUCTURE: Searches across all filters → file combos → entries
		/// </summary>
		/// <param name="doc">Revit document</param>
		/// <param name="categoryName">Category name</param>
		/// <param name="mepElementId">MEP element ID</param>
		/// <param name="structuralElementId">Structural element ID</param>
		/// <param name="intersectionPointX">Intersection point X coordinate</param>
		/// <param name="intersectionPointY">Intersection point Y coordinate</param>
		/// <param name="intersectionPointZ">Intersection point Z coordinate</param>
		/// <param name="tolerance">Tolerance for point matching (default: 0.1 feet)</param>
		/// <returns>Matching Global XML entry if found, null otherwise</returns>
		public static CategoryGlobalIndexEntry FindByMepHostAndPoint(Document doc, string categoryName, int mepElementId, int structuralElementId, double intersectionPointX, double intersectionPointY, double intersectionPointZ, double tolerance = 0.1)
		{
			if (mepElementId <= 0 || structuralElementId <= 0)
				return null;
			
			var index = LoadOrCreate(doc, categoryName);
			
			// ✅ HIERARCHICAL STRUCTURE: Get all entries from hierarchical + flat structures
			var allEntries = GetAllEntries(index);
			
			if (allEntries == null || !allEntries.Any())
				return null;
			
			// Round point to tolerance for comparison
			string pointKey = $"{Math.Round(intersectionPointX / tolerance) * tolerance:F1}," +
			                 $"{Math.Round(intersectionPointY / tolerance) * tolerance:F1}," +
			                 $"{Math.Round(intersectionPointZ / tolerance) * tolerance:F1}";
			
			CategoryGlobalIndexEntry bestMatch = null;
			double closestDistance = double.MaxValue;
			
			foreach (var entry in allEntries)
			{
				// First check: MEP and Host must match
				if (entry.MepElementId != mepElementId || entry.StructuralElementId != structuralElementId)
					continue;
				
				// Second check: Intersection point must be close (within tolerance)
				if (Math.Abs(entry.IntersectionPointX) > 1e-9 && 
				    Math.Abs(entry.IntersectionPointY) > 1e-9 && 
				    Math.Abs(entry.IntersectionPointZ) > 1e-9)
				{
					var existingPoint = new XYZ(entry.IntersectionPointX, entry.IntersectionPointY, entry.IntersectionPointZ);
					var newPoint = new XYZ(intersectionPointX, intersectionPointY, intersectionPointZ);
					var distance = existingPoint.DistanceTo(newPoint);
					
					if (distance <= tolerance)
					{
						// Found match - keep the closest one if multiple matches
						if (distance < closestDistance)
						{
							closestDistance = distance;
							bestMatch = entry;
						}
					}
				}
				else
				{
					// If existing entry has zero/invalid intersection point, still match by MEP+Host only
					// (fallback for old XML data without intersection points)
					if (bestMatch == null || closestDistance == double.MaxValue)
					{
						bestMatch = entry;
						closestDistance = 0; // Exact match by MEP+Host
					}
				}
			}
			
			return bestMatch;
		}

		/// <summary>
		/// ✅ CRITICAL: Builds a Dictionary keyed by (MEP+Host+Point) for O(1) lookup
		/// Used in RefreshService to quickly filter resolved intersections
		/// ✅ HIERARCHICAL STRUCTURE: Reads from all filters → file combos → entries
		/// </summary>
		public static Dictionary<(int mepId, int hostId, string pointKey), CategoryGlobalIndexEntry> BuildMepHostPointIndex(Document doc, string categoryName, double tolerance = 0.1)
		{
			var index = LoadOrCreate(doc, categoryName);
			var result = new Dictionary<(int mepId, int hostId, string pointKey), CategoryGlobalIndexEntry>();
			
			// ✅ HIERARCHICAL STRUCTURE: Get all entries from hierarchical + flat structures
			var allEntries = GetAllEntries(index);
			
			if (allEntries == null || !allEntries.Any())
				return result;
			
			foreach (var entry in allEntries)
			{
				// Only include entries with valid MEP+Host+Point data
				if (entry.MepElementId > 0 && entry.StructuralElementId > 0 &&
				    Math.Abs(entry.IntersectionPointX) > 1e-9 && 
				    Math.Abs(entry.IntersectionPointY) > 1e-9 && 
				    Math.Abs(entry.IntersectionPointZ) > 1e-9)
				{
					string pointKey = $"{Math.Round(entry.IntersectionPointX / tolerance) * tolerance:F1}," +
					                 $"{Math.Round(entry.IntersectionPointY / tolerance) * tolerance:F1}," +
					                 $"{Math.Round(entry.IntersectionPointZ / tolerance) * tolerance:F1}";
					
					var key = (entry.MepElementId, entry.StructuralElementId, pointKey);
					
					// If multiple entries with same key, prefer resolved ones
					if (!result.ContainsKey(key) || (!result[key].IsResolved && !result[key].IsClusterResolved))
					{
						result[key] = entry;
					}
				}
			}
			
			return result;
		}

		public static void Save(Document doc, CategoryGlobalIndex index)
		{
			// ✅ CRITICAL: Log EVERY Save() call to detect if file is being overwritten after flag reset
			string path = GetCategoryIndexPath(doc, index.Category);
			var entriesBeforeSave = GetAllEntries(index).ToList();
			int resolvedBeforeSave = entriesBeforeSave.Count(e => e.IsResolved);
			int clusterResolvedBeforeSave = entriesBeforeSave.Count(e => e.IsClusterResolved);
			DebugLogger.Info($"[GLOBAL-INDEX] 🔵 Save() CALLED for category '{index.Category}' at {DateTime.Now:HH:mm:ss.fff} - File: {path} - Memory: IsResolved={resolvedBeforeSave}, IsClusterResolved={clusterResolvedBeforeSave}");
			try
			{
				var refreshLogPath = Path.Combine(Path.GetDirectoryName(path) ?? "", "Refresh_GlobalIndex_Verification.log");
				SafeFileLogger.SafeAppendText(refreshLogPath, $"[{DateTime.Now}] [GLOBAL-INDEX] 🔵 Save() CALLED for category '{index.Category}' - File: {path} - Memory: IsResolved={resolvedBeforeSave}, IsClusterResolved={clusterResolvedBeforeSave}\n");
			}
			catch { }
			
			try
			{
				ProjectPathService.EnsureFiltersDirectory(doc);
				
				// ✅ CRITICAL FIX: Clean up placeholder FileComboGroups with empty LinkedFile/HostFile BEFORE saving
				// Remove FileComboGroups that have empty LinkedFile/HostFile (they shouldn't exist)
				if (index.Filters != null && index.Filters.Count > 0)
				{
					foreach (var filterGroup in index.Filters)
					{
						if (filterGroup.FileCombos != null && filterGroup.FileCombos.Count > 0)
						{
							// Remove placeholder FileComboGroups with empty LinkedFile/HostFile
							var placeholdersToRemove = filterGroup.FileCombos
								.Where(fc => string.IsNullOrWhiteSpace(fc.LinkedFile) || string.IsNullOrWhiteSpace(fc.HostFile))
								.ToList();
							
							foreach (var placeholder in placeholdersToRemove)
							{
								// Move entries from placeholder to flat structure (if needed)
								if (placeholder.Entries != null && placeholder.Entries.Count > 0)
								{
									if (index.Entries == null)
										index.Entries = new List<CategoryGlobalIndexEntry>();
									
									foreach (var entry in placeholder.Entries)
									{
										// Only add if not already in flat structure
										if (!index.Entries.Any(e => string.Equals(e.Id, entry.Id, StringComparison.OrdinalIgnoreCase)))
										{
											index.Entries.Add(entry);
										}
									}
								}
								
								filterGroup.FileCombos.Remove(placeholder);
							}
						}
					}
					
					// Remove empty FilterGroups (no FileCombos left)
					var emptyFilters = index.Filters.Where(f => f.FileCombos == null || f.FileCombos.Count == 0).ToList();
					foreach (var emptyFilter in emptyFilters)
					{
						index.Filters.Remove(emptyFilter);
					}
					
					// ✅ CRITICAL FIX: Remove FileComboGroups with NO ENTRIES (even if they have LinkedFile/HostFile)
					// These are placeholders created by MarkFileCombosAsProcessed but don't have actual clash zone entries
					foreach (var filterGroup in index.Filters)
					{
						if (filterGroup.FileCombos != null && filterGroup.FileCombos.Count > 0)
						{
							var emptyFileCombos = filterGroup.FileCombos
								.Where(fc => fc.Entries == null || fc.Entries.Count == 0)
								.ToList();
							
							foreach (var emptyFileCombo in emptyFileCombos)
							{
								filterGroup.FileCombos.Remove(emptyFileCombo);
							}
						}
					}
					
					// Remove FilterGroups that have no FileCombos left after cleanup
					var filtersToRemove = index.Filters.Where(f => f.FileCombos == null || f.FileCombos.Count == 0).ToList();
					foreach (var filterToRemove in filtersToRemove)
					{
						index.Filters.Remove(filterToRemove);
					}
				}
				
				// ✅ CRITICAL FIX: Remove ProcessedFileCombo entries that are already represented in FilterGroups
				// This prevents duplicate file combo information at root level
				if (index.ProcessedFileCombos != null && index.ProcessedFileCombos.Count > 0 && index.Filters != null && index.Filters.Count > 0)
				{
					var hierarchicalFileComboKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
					
					// Collect all file combo keys from hierarchical structure
					foreach (var filterGroup in index.Filters)
					{
						if (filterGroup.FileCombos != null)
						{
							foreach (var fileCombo in filterGroup.FileCombos)
							{
								if (!string.IsNullOrWhiteSpace(fileCombo.LinkedFile) && !string.IsNullOrWhiteSpace(fileCombo.HostFile))
								{
									var combo = new ProcessedFileCombo { LinkedFile = fileCombo.LinkedFile, HostFile = fileCombo.HostFile };
									hierarchicalFileComboKeys.Add(combo.GetNormalizedKey());
								}
							}
						}
					}
					
					// Remove ProcessedFileCombo entries that are already in hierarchical structure
					if (hierarchicalFileComboKeys.Count > 0)
					{
						index.ProcessedFileCombos = index.ProcessedFileCombos
							.Where(pfc => !hierarchicalFileComboKeys.Contains(pfc.GetNormalizedKey()))
							.ToList();
					}
				}
				
				// ✅ CRITICAL FIX: Remove duplicate entries from flat structure that are already in hierarchical structure
				// This prevents serializing the same entry twice
				if (index.Entries != null && index.Entries.Count > 0 && index.Filters != null && index.Filters.Count > 0)
				{
					var hierarchicalEntryIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
					
					// Collect all entry IDs from hierarchical structure
					foreach (var filterGroup in index.Filters)
					{
						if (filterGroup.FileCombos != null)
						{
							foreach (var fileCombo in filterGroup.FileCombos)
							{
								if (fileCombo.Entries != null)
								{
									foreach (var entry in fileCombo.Entries)
									{
										if (!string.IsNullOrEmpty(entry.Id))
											hierarchicalEntryIds.Add(entry.Id);
									}
								}
							}
						}
					}
					
					// Remove duplicates from flat structure
					if (hierarchicalEntryIds.Count > 0)
					{
						index.Entries = index.Entries
							.Where(e => !hierarchicalEntryIds.Contains(e.Id))
							.ToList();
					}
				}
				
				// ✅ CRITICAL FIX: Ensure ProcessedFileCombos is initialized before saving
				if (index.ProcessedFileCombos == null)
					index.ProcessedFileCombos = new List<ProcessedFileCombo>();
				
				// ✅ CRITICAL: Log before saving to verify what we're about to write
				var allEntriesBeforeSave = GetAllEntries(index).ToList();
				int resolvedCount = allEntriesBeforeSave.Count(e => e.IsResolved);
				int clusterResolvedCount = allEntriesBeforeSave.Count(e => e.IsClusterResolved);
				DebugLogger.Info($"[GLOBAL-INDEX] 📝 About to save Global XML for '{index.Category}': Total entries={allEntriesBeforeSave.Count}, IsResolved={resolvedCount}, IsClusterResolved={clusterResolvedCount}");
				
				// ✅ CRITICAL: Ensure file is written and flushed to disk
				using (var writer = new StreamWriter(path))
				{
					_serializer.Serialize(writer, index);
					writer.Flush(); // ✅ Force flush to ensure data is written to disk
				}
				
				// ✅ CRITICAL: Small delay to ensure file system has written the file
				System.Threading.Thread.Sleep(50);
				
				// ✅ VERIFICATION: Verify file was written and check flag values
				if (File.Exists(path))
				{
					var fileContent = File.ReadAllText(path);
					bool hasProcessedFileCombos = fileContent.Contains("<ProcessedFileCombo");
					
					// ✅ CRITICAL: Verify the saved file contains the updated flag values
					int trueResolvedInFile = (fileContent.Split(new[] { "IsResolved=\"true\"" }, StringSplitOptions.None).Length - 1);
					int falseResolvedInFile = (fileContent.Split(new[] { "IsResolved=\"false\"" }, StringSplitOptions.None).Length - 1);
					int trueClusterInFile = (fileContent.Split(new[] { "IsClusterResolved=\"true\"" }, StringSplitOptions.None).Length - 1);
					int falseClusterInFile = (fileContent.Split(new[] { "IsClusterResolved=\"false\"" }, StringSplitOptions.None).Length - 1);
					
					// ✅ CRITICAL: Always log to both DebugLogger AND direct file (bypasses filtering)
					string verificationMsg = $"[GLOBAL-INDEX] ✅ Saved Global XML for '{index.Category}' to: {path}";
					string fileVerificationMsg = $"[GLOBAL-INDEX] 📊 File verification: IsResolved true={trueResolvedInFile}, false={falseResolvedInFile} | IsClusterResolved true={trueClusterInFile}, false={falseClusterInFile}";
					string memoryStateMsg = $"[GLOBAL-INDEX] 📊 Memory state: IsResolved={resolvedCount}, IsClusterResolved={clusterResolvedCount}";
					
					DebugLogger.Info(verificationMsg);
					DebugLogger.Info(fileVerificationMsg);
					DebugLogger.Info(memoryStateMsg);
					
					// ✅ DIRECT FILE LOGGING: Write to refresh log file to bypass DebugLogger filtering
					try
					{
						var refreshLogPath = Path.Combine(Path.GetDirectoryName(path) ?? "", "Refresh_GlobalIndex_Verification.log");
						SafeFileLogger.SafeAppendText(refreshLogPath, $"[{DateTime.Now}] {verificationMsg}\n");
						SafeFileLogger.SafeAppendText(refreshLogPath, $"[{DateTime.Now}] {fileVerificationMsg}\n");
						SafeFileLogger.SafeAppendText(refreshLogPath, $"[{DateTime.Now}] {memoryStateMsg}\n");
					}
					catch { /* Ignore file logging errors */ }
					
					if (trueResolvedInFile != resolvedCount || trueClusterInFile != clusterResolvedCount)
					{
						string mismatchMsg = $"[GLOBAL-INDEX] ⚠️ MISMATCH: File flags don't match memory state! File has {trueResolvedInFile} IsResolved=true, {trueClusterInFile} IsClusterResolved=true, but memory has {resolvedCount} IsResolved, {clusterResolvedCount} IsClusterResolved";
						DebugLogger.Warning(mismatchMsg);
						try
						{
							var refreshLogPath = Path.Combine(Path.GetDirectoryName(path) ?? "", "Refresh_GlobalIndex_Verification.log");
							SafeFileLogger.SafeAppendText(refreshLogPath, $"[{DateTime.Now}] {mismatchMsg}\n");
						}
						catch { /* Ignore file logging errors */ }
					}
				}
				else
				{
					string errorMsg = $"[GLOBAL-INDEX] ❌ File was NOT created at: {path}";
					DebugLogger.Error(errorMsg);
					try
					{
						var refreshLogPath = Path.Combine(Path.GetDirectoryName(path) ?? "", "Refresh_GlobalIndex_Verification.log");
						SafeFileLogger.SafeAppendText(refreshLogPath, $"[{DateTime.Now}] {errorMsg}\n");
					}
					catch { /* Ignore file logging errors */ }
				}
			}
			catch (Exception ex)
			{
				if (!DeploymentConfiguration.DeploymentMode)
					DebugLogger.Error($"[GLOBAL-INDEX] ERROR saving Global XML for '{index.Category}': {ex.Message}\n{ex.StackTrace}");
				SafeFileLogger.SafeAppendText("Refresh_GlobalIndex.log", $"Failed to save global index for '{index.Category}': {ex.Message}\n{ex.StackTrace}");
			}
		}

		/// <summary>
		/// ✅ OPTIMIZATION: Checks if a file combination (LinkedFile + HostFile) has been processed for a category
		/// Returns true if the combo was already processed in ANY filter, false if it's new
		/// ✅ GLOBAL CHECK: Searches across all filters in hierarchical structure
		/// </summary>
		/// <param name="doc">Revit document</param>
		/// <param name="categoryName">Category name</param>
		/// <param name="linkedFile">Linked file name (reference file)</param>
		/// <param name="hostFile">Host file name (structural file)</param>
		/// <returns>True if file combo was already processed in any filter, false if new</returns>
		public static bool IsFileComboProcessed(Document doc, string categoryName, string linkedFile, string hostFile)
		{
			if (string.IsNullOrWhiteSpace(categoryName) || string.IsNullOrWhiteSpace(linkedFile) || string.IsNullOrWhiteSpace(hostFile))
				return false;
			
			try
			{
				var index = LoadOrCreate(doc, categoryName);
				var combo = new ProcessedFileCombo { LinkedFile = linkedFile, HostFile = hostFile };
				var normalizedKey = combo.GetNormalizedKey();
				
				// ✅ HIERARCHICAL STRUCTURE: Check all filters for processed combo
				if (index.Filters != null && index.Filters.Count > 0)
				{
					foreach (var filter in index.Filters)
					{
						if (filter.FileCombos != null)
						{
							foreach (var fileCombo in filter.FileCombos)
							{
								if (fileCombo.IsProcessed && fileCombo.GetNormalizedKey() == normalizedKey)
									return true; // Found processed combo in any filter
							}
						}
					}
				}
				
				// ✅ FALLBACK: Check flat structure (backward compatibility)
				if (index.ProcessedFileCombos != null && index.ProcessedFileCombos.Count > 0)
				{
					if (index.ProcessedFileCombos.Any(pfc => pfc.GetNormalizedKey() == normalizedKey))
						return true;
				}
				
				return false;
			}
			catch (Exception ex)
			{
				SafeFileLogger.SafeAppendText("Refresh_GlobalIndex.log", 
					$"Error checking processed file combo for '{categoryName}': {ex.Message}");
				return false; // On error, treat as new to be safe
			}
		}
		
		/// <summary>
		/// ✅ OPTIMIZATION: Marks a file combination (LinkedFile + HostFile) as processed for a category
		/// Called after intersection detection completes to track which file combos have been processed
		/// </summary>
		/// <param name="doc">Revit document</param>
		/// <param name="categoryName">Category name</param>
		/// <param name="linkedFile">Linked file name (reference file)</param>
		/// <param name="hostFile">Host file name (structural file)</param>
		public static void MarkFileComboAsProcessed(Document doc, string categoryName, string linkedFile, string hostFile)
		{
			if (string.IsNullOrWhiteSpace(categoryName) || string.IsNullOrWhiteSpace(linkedFile) || string.IsNullOrWhiteSpace(hostFile))
				return;
			
			try
			{
				var index = LoadOrCreate(doc, categoryName);
				if (index.ProcessedFileCombos == null)
					index.ProcessedFileCombos = new List<ProcessedFileCombo>();
				
				var combo = new ProcessedFileCombo { LinkedFile = linkedFile, HostFile = hostFile, ProcessedAt = DateTime.Now };
				var normalizedKey = combo.GetNormalizedKey();
				
				// Check if combo already exists (avoid duplicates)
				bool exists = index.ProcessedFileCombos.Any(pfc => pfc.GetNormalizedKey() == normalizedKey);
				
				if (!exists)
				{
					index.ProcessedFileCombos.Add(combo);
					Save(doc, index);
					
					if (!DeploymentConfiguration.DeploymentMode)
						DebugLogger.Info($"[GLOBAL-INDEX] Marked file combo as processed: Category='{categoryName}', LinkedFile='{linkedFile}', HostFile='{hostFile}'");
				}
			}
			catch (Exception ex)
			{
				SafeFileLogger.SafeAppendText("Refresh_GlobalIndex.log", 
					$"Error marking file combo as processed for '{categoryName}': {ex.Message}");
			}
		}
		
		/// <summary>
		/// ✅ OPTIMIZATION: Marks multiple file combinations as processed (batch operation)
		/// ✅ HIERARCHICAL STRUCTURE: Writes to Filter → FileCombo in hierarchical structure
		/// </summary>
		/// <param name="doc">Revit document</param>
		/// <param name="categoryName">Category name</param>
		/// <param name="fileCombos">List of (LinkedFile, HostFile) tuples to mark as processed</param>
		/// <param name="filterName">Filter name (required for hierarchical structure)</param>
		public static void MarkFileCombosAsProcessed(Document doc, string categoryName, IEnumerable<(string LinkedFile, string HostFile)> fileCombos, string filterName = null)
		{
			if (string.IsNullOrWhiteSpace(categoryName) || fileCombos == null)
				return;
			
			try
			{
				var index = LoadOrCreate(doc, categoryName);
				
				// ✅ HIERARCHICAL STRUCTURE: Ensure Filters list is initialized
				if (index.Filters == null)
					index.Filters = new List<FilterGroup>();
				
				// ✅ HIERARCHICAL STRUCTURE: Get or create FilterGroup for this filter
				var filterGroup = index.Filters.FirstOrDefault(f => f.Name == filterName);
				if (filterGroup == null && !string.IsNullOrEmpty(filterName))
				{
					filterGroup = new FilterGroup { Name = filterName, FileCombos = new List<FileComboGroup>() };
					index.Filters.Add(filterGroup);
				}
				
				// ✅ FALLBACK: Keep flat structure for backward compatibility (deprecated)
				if (index.ProcessedFileCombos == null)
					index.ProcessedFileCombos = new List<ProcessedFileCombo>();
				
				var existingKeys = new HashSet<string>(index.ProcessedFileCombos.Select(pfc => pfc.GetNormalizedKey()), StringComparer.OrdinalIgnoreCase);
				bool changed = false;
				int addedCount = 0;
				int existingCount = 0;
				int hierarchicalAddedCount = 0;
				
				foreach (var (linkedFile, hostFile) in fileCombos)
				{
					if (string.IsNullOrWhiteSpace(linkedFile) || string.IsNullOrWhiteSpace(hostFile))
						continue;
					
					var combo = new ProcessedFileCombo { LinkedFile = linkedFile, HostFile = hostFile, ProcessedAt = DateTime.Now, IsProcessed = true };
					var normalizedKey = combo.GetNormalizedKey();
					
					// ✅ HIERARCHICAL STRUCTURE: Add to FilterGroup → FileComboGroup
					if (filterGroup != null && !string.IsNullOrEmpty(filterName))
					{
						if (filterGroup.FileCombos == null)
							filterGroup.FileCombos = new List<FileComboGroup>();
						
						var fileComboGroup = filterGroup.FileCombos.FirstOrDefault(fc => fc.GetNormalizedKey() == normalizedKey);
						if (fileComboGroup == null)
						{
							fileComboGroup = new FileComboGroup
							{
								LinkedFile = linkedFile,
								HostFile = hostFile,
								ProcessedAt = DateTime.Now,
								IsProcessed = true,
								Entries = new List<CategoryGlobalIndexEntry>()
							};
							filterGroup.FileCombos.Add(fileComboGroup);
							hierarchicalAddedCount++;
							changed = true;
						}
						else if (!fileComboGroup.IsProcessed)
						{
							fileComboGroup.IsProcessed = true;
							fileComboGroup.ProcessedAt = DateTime.Now;
							changed = true;
						}
					}
					
					// ✅ FALLBACK: Also update flat structure for backward compatibility
					if (!existingKeys.Contains(normalizedKey))
					{
						index.ProcessedFileCombos.Add(combo);
						existingKeys.Add(normalizedKey);
						changed = true;
						addedCount++;
					}
					else
					{
						var existingCombo = index.ProcessedFileCombos.FirstOrDefault(pfc => pfc.GetNormalizedKey().Equals(normalizedKey, StringComparison.OrdinalIgnoreCase));
						if (existingCombo != null && !existingCombo.IsProcessed)
						{
							existingCombo.IsProcessed = true;
							existingCombo.ProcessedAt = DateTime.Now;
							changed = true;
						}
						existingCount++;
					}
				}
				
				// ✅ CRITICAL FIX: ALWAYS save when MarkFileCombosAsProcessed is called with valid fileCombos
				if (fileCombos.Any())
				{
					Save(doc, index);
					if (!DeploymentConfiguration.DeploymentMode)
					{
						DebugLogger.Info($"[GLOBAL-INDEX] ✅ Saved Global XML for '{categoryName}' - Filter '{filterName ?? "Unknown"}' - {hierarchicalAddedCount} file combos added to hierarchical structure");
						DebugLogger.Info($"[GLOBAL-INDEX]   Flat structure: {index.ProcessedFileCombos.Count} ProcessedFileCombos total ({addedCount} new added, {existingCount} already existed)");
						DebugLogger.Info($"[GLOBAL-INDEX]   File path: {GetCategoryIndexPath(doc, categoryName)}");
					}
				}
				else if (!DeploymentConfiguration.DeploymentMode)
				{
					DebugLogger.Warning($"[GLOBAL-INDEX] ⚠️ No file combos to save for '{categoryName}' (fileCombos is empty or null)");
				}
			}
			catch (Exception ex)
			{
				if (!DeploymentConfiguration.DeploymentMode)
					DebugLogger.Error($"[GLOBAL-INDEX] ERROR marking file combos for '{categoryName}': {ex.Message}\n{ex.StackTrace}");
				SafeFileLogger.SafeAppendText("Refresh_GlobalIndex.log", 
					$"Error marking file combos as processed for '{categoryName}': {ex.Message}\n{ex.StackTrace}");
			}
		}
		
		/// <summary>
		/// ✅ OPTIMIZATION: Gets all processed file combinations for a category
		/// Returns list of normalized file combo keys
		/// </summary>
		/// <param name="doc">Revit document</param>
		/// <param name="categoryName">Category name</param>
		/// <returns>Set of normalized file combo keys (format: "linkedfile|hostfile")</returns>
	public static HashSet<string> GetProcessedFileComboKeys(Document doc, string categoryName)
	{
		var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
		
		if (string.IsNullOrWhiteSpace(categoryName))
			return keys;
		
		try
		{
			var index = LoadOrCreate(doc, categoryName);
			
			// ✅ FIX: Check hierarchical structure first (Filters → FileCombos)
			// This matches the logic in IsFileComboProcessed
			if (index?.Filters != null && index.Filters.Count > 0)
			{
				foreach (var filter in index.Filters)
				{
					if (filter.FileCombos != null)
					{
						foreach (var fileCombo in filter.FileCombos)
						{
							// ✅ Only include combos that are marked as processed
							if (fileCombo.IsProcessed)
							{
								var key = fileCombo.GetNormalizedKey();
								if (!string.IsNullOrEmpty(key))
									keys.Add(key);
							}
						}
					}
				}
			}
			
			// ✅ FALLBACK: Also check flat structure (backward compatibility)
			if (index?.ProcessedFileCombos != null && index.ProcessedFileCombos.Count > 0)
			{
				foreach (var combo in index.ProcessedFileCombos)
				{
					var key = combo.GetNormalizedKey();
					if (!string.IsNullOrEmpty(key))
						keys.Add(key);
				}
			}
		}
		catch (Exception ex)
		{
			SafeFileLogger.SafeAppendText("Refresh_GlobalIndex.log", 
				$"Error getting processed file combo keys for '{categoryName}': {ex.Message}");
		}
		
		return keys;
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
