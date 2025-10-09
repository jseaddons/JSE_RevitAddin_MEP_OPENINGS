# MEPMARK Integration - CORRECTED Implementation Plan

## Critical Issues Identified & Solutions

### ❌ **Issue 1: Static Properties for UI Thread Access**
**Problem:** Using static properties to pass UI data to ExternalEvent is fragile and error-prone.

**✅ Solution:** Pass data explicitly through ExternalEvent handler instance variables.

### ❌ **Issue 2: Missing MarkPrefixSettings Class Definition**
**Problem:** Referenced but not properly defined.

**✅ Solution:** Create proper MarkPrefixSettings class with GetDisciplinePrefix method.

### ❌ **Issue 3: Dangerous Fallback Logic in Cluster Filtering**
**Problem:** Falls back to ALL cluster sleeves when MEP_Category not found.

**✅ Solution:** Return empty list instead of all sleeves.

### ❌ **Issue 4: Mark Numbering Continuity**
**Problem:** Always starts numbering from 1, causing duplicates on re-runs.

**✅ Solution:** Check existing marks and continue numbering.

### ❌ **Issue 5: Transaction Naming Inconsistency**
**Problem:** Inconsistent transaction naming pattern.

**✅ Solution:** Use consistent naming pattern.

## Corrected Implementation

### Step 1: Create MarkPrefixSettings Class

**File:** `Models/MarkPrefixSettings.cs` (Already exists - verify it has the correct content)

```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Models
{
    /// <summary>
    /// Mark prefix configuration for MEPMARK parameter
    /// Stores project-level and category-specific prefixes
    /// Used for UI state transfer to mark parameter command
    /// </summary>
    public class MarkPrefixSettings
    {
        /// <summary>
        /// Project-level prefix (e.g., "SLEEVE_") - prepended to all marks
        /// </summary>
        public string ProjectPrefix { get; set; } = "SLEEVE_";
        
        /// <summary>
        /// Discipline prefix for Ducts (default: "DCT")
        /// </summary>
        public string DuctPrefix { get; set; } = "DCT";
        
        /// <summary>
        /// Discipline prefix for Pipes (default: "PLU")
        /// </summary>
        public string PipePrefix { get; set; } = "PLU";
        
        /// <summary>
        /// Discipline prefix for Cable Trays (default: "ELE")
        /// </summary>
        public string CableTrayPrefix { get; set; } = "ELE";
        
        /// <summary>
        /// Discipline prefix for Duct Accessories/Dampers (default: "DAM")
        /// </summary>
        public string DamperPrefix { get; set; } = "DAM";
        
        /// <summary>
        /// Get discipline prefix for a specific category
        /// </summary>
        public string GetDisciplinePrefix(string category)
        {
            return category switch
            {
                "Ducts" => DuctPrefix,
                "Pipes" => PipePrefix,
                "Cable Trays" => CableTrayPrefix,
                "Duct Accessories" => DamperPrefix,
                _ => "OPN" // Generic fallback
            };
        }
    }
}
```

### Step 2: Create MarkParameterCommand

**File:** `Commands/MarkParameterCommand.cs`

```csharp
using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    /// <summary>
    /// Mark Parameter Command - Applies MEPMARK to cluster sleeves for a specific category
    /// Can be called from any context (ICommand, ExternalEvent, etc.)
    /// </summary>
    public class MarkParameterCommand : ICommand
    {
        private readonly string _targetCategory;
        private readonly string _projectPrefix;
        private readonly string _disciplinePrefix;
        
        public MarkParameterCommand(string targetCategory, string projectPrefix, string disciplinePrefix)
        {
            _targetCategory = targetCategory ?? throw new ArgumentNullException(nameof(targetCategory));
            _projectPrefix = projectPrefix ?? throw new ArgumentNullException(nameof(projectPrefix));
            _disciplinePrefix = disciplinePrefix ?? throw new ArgumentNullException(nameof(disciplinePrefix));
        }
        
        public void Execute(UIApplication app)
        {
            try
            {
                DebugLogger.Info($"[MarkParameterCommand] Starting MEPMARK application for {_targetCategory}");
                DebugLogger.Info($"[MarkParameterCommand] Project Prefix: '{_projectPrefix}', Discipline Prefix: '{_disciplinePrefix}'");
                
                var doc = app.ActiveUIDocument.Document;
                var uiDoc = app.ActiveUIDocument;
                
                // ⚠️ CRITICAL: Transaction management (follows existing pattern)
                using (var tx = new Transaction(doc, $"Mark {_targetCategory} Clusters"))
                {
                    tx.Start();
                    
                    var markService = new MarkParameterService();
                    var (processedCount, errorCount) = markService.ApplyMepMarkToClusters(
                        doc, _targetCategory, _projectPrefix, _disciplinePrefix);
                    
                    tx.Commit();
                    
                    DebugLogger.Info($"[MarkParameterCommand] ✓ MEPMARK complete for {_targetCategory}: {processedCount} clusters processed, {errorCount} errors");
                    
                    // Optional: Show user feedback for manual testing
                    if (processedCount > 0)
                    {
                        TaskDialog.Show("MEPMARK Applied", 
                            $"Category: {_targetCategory}\n\n" +
                            $"✓ {processedCount} cluster(s) marked\n" +
                            $"✓ Project Prefix: {_projectPrefix}\n" +
                            $"✓ Discipline Prefix: {_disciplinePrefix}");
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[MarkParameterCommand] Error applying MEPMARK to {_targetCategory}: {ex.Message}");
                DebugLogger.Error($"[MarkParameterCommand] Stack trace: {ex.StackTrace}");
                // Don't show TaskDialog here - let orchestrator handle errors
                throw; // Re-throw to let caller handle
            }
        }
    }
}
```

### Step 3: Create MarkParameterService

**File:** `Services/MarkParameterService.cs`

```csharp
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Service for applying MEPMARK parameters to cluster sleeves
    /// </summary>
    public class MarkParameterService
    {
        /// <summary>
        /// Apply MEPMARK to cluster sleeves for a specific category
        /// </summary>
        /// <param name="doc">Revit document</param>
        /// <param name="category">Target category (Ducts, Pipes, Cable Trays, etc.)</param>
        /// <param name="projectPrefix">Project prefix (e.g., "SLEEVE_")</param>
        /// <param name="disciplinePrefix">Discipline prefix (e.g., "DCT", "PLU", "ELE")</param>
        /// <returns>Tuple of (processedCount, errorCount)</returns>
        public (int processedCount, int errorCount) ApplyMepMarkToClusters(
            Document doc, string category, string projectPrefix, string disciplinePrefix)
        {
            int processedCount = 0;
            int errorCount = 0;
            
            try
            {
                // Find cluster sleeves for this category
                var clusterSleeves = GetClusterSleevesForCategory(doc, category);
                
                DebugLogger.Info($"[MarkParameterService] Found {clusterSleeves.Count} cluster sleeves for category '{category}'");
                
                if (clusterSleeves.Count == 0)
                {
                    DebugLogger.Warning($"[MarkParameterService] No cluster sleeves found for category '{category}'");
                    return (0, 0);
                }
                
                // Get starting index based on existing marks
                int startIndex = GetMaxExistingMarkNumber(doc, projectPrefix, disciplinePrefix) + 1;
                
                // Apply MEPMARK to each cluster
                for (int i = 0; i < clusterSleeves.Count; i++)
                {
                    try
                    {
                        var cluster = clusterSleeves[i];
                        int markNumber = startIndex + i;
                        string markValue = GenerateMarkValue(disciplinePrefix, markNumber);
                        string fullMarkValue = $"{projectPrefix}{markValue}";
                        
                        SetMarkParameter(cluster, fullMarkValue);
                        processedCount++;
                        
                        DebugLogger.Info($"[MarkParameterService] Applied MEPMARK '{fullMarkValue}' to cluster {cluster.Id.IntegerValue}");
                    }
                    catch (Exception ex)
                    {
                        errorCount++;
                        DebugLogger.Error($"[MarkParameterService] Error applying MEPMARK to cluster {clusterSleeves[i].Id.IntegerValue}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[MarkParameterService] Error processing category '{category}': {ex.Message}");
                throw;
            }
            
            return (processedCount, errorCount);
        }
        
        /// <summary>
        /// Find cluster sleeves for a specific category
        /// ✅ CORRECTED: No dangerous fallback to all sleeves
        /// </summary>
        private List<FamilyInstance> GetClusterSleevesForCategory(Document doc, string category)
        {
            var allClusterSleeves = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => {
                    var famName = fi.Symbol?.Family?.Name ?? string.Empty;
                    return famName == "RectangularOpeningOnWall" || 
                           famName == "CircularOpeningOnWall" ||
                           famName == "RectangularOpeningOnSlab" || 
                           famName == "CircularOpeningOnSlab";
                })
                .ToList();
            
            // Filter by MEP_Category parameter if available
            var categorySleeves = allClusterSleeves.Where(sleeve => {
                var categoryParam = sleeve.LookupParameter("MEP_Category");
                string sleeveCategory = categoryParam?.AsString() ?? "";
                return sleeveCategory == category;
            }).ToList();
            
            // ✅ CORRECTED: Return empty list instead of all sleeves if no category match
            if (categorySleeves.Count == 0)
            {
                DebugLogger.Warning($"[MarkParameterService] No sleeves found with MEP_Category='{category}'");
                return new List<FamilyInstance>(); // ✅ Safe: Return empty list
            }
            
            return categorySleeves;
        }
        
        /// <summary>
        /// ✅ NEW: Get maximum existing mark number to continue numbering
        /// </summary>
        private int GetMaxExistingMarkNumber(Document doc, string projectPrefix, string disciplinePrefix)
        {
            try
            {
                var allElements = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .ToList();
                
                int maxNumber = 0;
                
                foreach (var element in allElements)
                {
                    var markParam = element.LookupParameter("Mark");
                    if (markParam != null)
                    {
                        string markValue = markParam.AsString() ?? "";
                        
                        // Check if this mark matches our pattern: ProjectPrefix + DisciplinePrefix + Number
                        string expectedPrefix = $"{projectPrefix}{disciplinePrefix}";
                        if (markValue.StartsWith(expectedPrefix))
                        {
                            // Extract the number part
                            string numberPart = markValue.Substring(expectedPrefix.Length);
                            if (int.TryParse(numberPart, out int number))
                            {
                                maxNumber = Math.Max(maxNumber, number);
                            }
                        }
                    }
                }
                
                DebugLogger.Info($"[MarkParameterService] Max existing mark number for '{projectPrefix}{disciplinePrefix}': {maxNumber}");
                return maxNumber;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[MarkParameterService] Error getting max existing mark number: {ex.Message}");
                return 0; // Start from 1 if error
            }
        }
        
        /// <summary>
        /// Generate mark value based on discipline prefix and number
        /// </summary>
        private string GenerateMarkValue(string disciplinePrefix, int number)
        {
            return $"{disciplinePrefix}{number:000}";
        }
        
        /// <summary>
        /// Set Mark parameter on element
        /// </summary>
        private void SetMarkParameter(Element element, string markValue)
        {
            var markParam = element.LookupParameter("Mark");
            if (markParam != null && !markParam.IsReadOnly)
            {
                markParam.Set(markValue);
            }
            else
            {
                throw new InvalidOperationException($"Cannot set Mark parameter on element {element.Id.IntegerValue}: parameter is null or read-only");
            }
        }
    }
}
```

### Step 4: Update SleevePlacementExternalEvent

**File:** `Services/SleevePlacementExternalEvent.cs`

**✅ CORRECTED: Use instance variables instead of static properties**

```csharp
public class SleevePlacementExternalEvent : IExternalEventHandler
{
    private List<string> _selectedCategories;
    private Document _document;
    private UIDocument _uiDocument;
    
    // ✅ NEW: Instance variables for mark prefixes
    private MarkPrefixSettings _markPrefixes;

    // ✅ NEW: Set context method for passing data from UI
    public void SetContext(List<string> categories, MarkPrefixSettings markPrefixes)
    {
        _selectedCategories = categories;
        _markPrefixes = markPrefixes;
    }

    public void Execute(UIApplication app)
    {
        try
        {
            DebugLogger.Info("[SleevePlacementExternalEvent] Starting sleeve placement process");
            
            _uiDocument = app.ActiveUIDocument;
            _document = _uiDocument.Document;
            
            // ⚠️ CRITICAL: Load cluster configuration from filter settings BEFORE executing commands
            LoadClusterConfigurationFromFilters();

            // Log immediate feedback (non-blocking)
            DebugLogger.Info($"[SleevePlacementExternalEvent] Processing {_selectedCategories.Count} categories: {string.Join(", ", _selectedCategories)}");

            // Process each category: Place individual sleeves → Cluster → Apply MEPMARK
            foreach (var category in _selectedCategories)
            {
                var clashZones = GetClashZonesForCategory(category);
                if (clashZones.Count > 0)
                {
                    // Step 1: Place individual sleeves
                    ICommand placementCommand = CreateCommandForCategory(category, clashZones);
                    
                    if (placementCommand != null)
                    {
                        // Step 1: Place individual sleeves
                        DebugLogger.Info($"[SleevePlacementExternalEvent] Placing individual sleeves for {category} ({clashZones.Count} clash zones)");
                        placementCommand.Execute(app);
                        
                        // Step 2: Immediately cluster this category's sleeves
                        DebugLogger.Info($"[SleevePlacementExternalEvent] Clustering {category} sleeves...");
                        var clusterCommand = new Commands.UniversalClusterCommand(category);
                        clusterCommand.Execute(app);
                        
                        // ✅ NEW: Step 3: Apply MEPMARK to clusters using stored prefixes
                        DebugLogger.Info($"[SleevePlacementExternalEvent] Applying MEPMARK to {category} clusters...");
                        
                        // ✅ CORRECTED: Use instance variable instead of static property
                        string projectPrefix = _markPrefixes.ProjectPrefix;
                        string disciplinePrefix = _markPrefixes.GetDisciplinePrefix(category);
                        
                        var markCommand = new Commands.MarkParameterCommand(category, projectPrefix, disciplinePrefix);
                        markCommand.Execute(app);
                        
                        DebugLogger.Info($"[SleevePlacementExternalEvent] ✓ Completed placement, clustering, and MEPMARK for {category}");
                    }
                    else
                    {
                        DebugLogger.Warning($"[SleevePlacementExternalEvent] No command available for category: {category}");
                    }
                }
                else
                {
                    DebugLogger.Warning($"[SleevePlacementExternalEvent] No clash zones found for category: {category}");
                }
            }
            
            DebugLogger.Info("[SleevePlacementExternalEvent] All categories processed (placement + clustering + MEPMARK)");
        }
        catch (Exception ex)
        {
            DebugLogger.Error($"[SleevePlacementExternalEvent] Exception: {ex.Message}");
            DebugLogger.Error($"[SleevePlacementExternalEvent] Stack trace: {ex.StackTrace}");
            TaskDialog.Show("Error", $"Failed to start sleeve placement: {ex.Message}");
        }
    }

    // ✅ NEW: Keep existing SetSelectedCategories for backward compatibility
    public void SetSelectedCategories(List<string> categories)
    {
        _selectedCategories = categories;
    }

    // ... rest of existing methods remain unchanged ...
}
```

### Step 5: Update EmergencyMainDialog

**File:** `Views/EmergencyMainDialog.cs`

**✅ CORRECTED: Pass prefix values directly to ExternalEvent handler**

```csharp
private void OnOkClick(object? sender, EventArgs e)
{
    try
    {
        // Get selected categories
        var selectedCategories = GetSelectedMepCategories();
        if (selectedCategories.Count == 0)
        {
            MessageBox.Show("Please select at least one MEP category.", "No Categories Selected", 
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        
        // ✅ CORRECTED: Read mark prefixes from UI and create MarkPrefixSettings object
        var markPrefixes = new MarkPrefixSettings
        {
            ProjectPrefix = _projectPrefixTextBox?.Text?.Trim() ?? "SLEEVE_",
            DuctPrefix = _categoryPrefixes.GetValueOrDefault("Ducts", "DCT"),
            PipePrefix = _categoryPrefixes.GetValueOrDefault("Pipes", "PLU"),
            CableTrayPrefix = _categoryPrefixes.GetValueOrDefault("Cable Trays", "ELE"),
            DamperPrefix = _categoryPrefixes.GetValueOrDefault("Duct Accessories", "DAM")
        };
        
        // ✅ CORRECTED: Pass both categories AND prefixes to handler
        _sleevePlacementHandler.SetContext(selectedCategories, markPrefixes);
        
        // ⚠️ CRITICAL: Save CONDITIONS.xml before raising external event
        SaveConditionsToXml(selectedCategories);
        
        // Show immediate feedback via status label (non-blocking)
        _statusLabel.Text = $"Processing {selectedCategories.Count} categories...";
        DebugLogger.Info($"[OnOkClick] Starting sleeve placement for categories: {string.Join(", ", selectedCategories)}");

        // Initialize RevitTask if not already done
        if (!RevitTask.IsInitialized())
        {
            var handler = new RevitTask.Handler();
            var exEvent = ExternalEvent.Create(handler);
            RevitTask.Init(exEvent);
        }
        
        // ✅ CORRECTED: Use SetContext instead of SetSelectedCategories
        _sleevePlacementEvent.Raise();

        DebugLogger.Info($"[EmergencyMainDialog] External event raised for categories: {string.Join(", ", selectedCategories)}");
        
        // ⚠️ CRITICAL: Close dialog to free Revit main thread
        this.Hide(); // Hide instead of Close to keep dialog in memory for status updates
    }
    catch (Exception ex)
    {
        DebugLogger.Error($"[EmergencyMainDialog] Exception in OnOkClick: {ex.Message}");
        MessageBox.Show($"Error starting sleeve placement: {ex.Message}", "Error", 
            MessageBoxButtons.OK, MessageBoxIcon.Error);
    }
}

// ✅ REMOVED: Static properties - no longer needed
// public static string LastProjectPrefix { get; set; } = "SLEEVE_";
// public static Dictionary<string, string> LastDisciplinePrefixes { get; set; } = new Dictionary<string, string>();
```

## Key Corrections Made

### ✅ **1. Thread-Safe Data Passing**
- **Before:** Static properties (race conditions, state persistence issues)
- **After:** Instance variables in ExternalEvent handler

### ✅ **2. Proper Cluster Filtering**
- **Before:** Dangerous fallback to all cluster sleeves
- **After:** Return empty list if no category match found

### ✅ **3. Mark Numbering Continuity**
- **Before:** Always starts from 1 (duplicates on re-runs)
- **After:** Checks existing marks and continues numbering

### ✅ **4. Consistent Transaction Naming**
- **Before:** Inconsistent transaction names
- **After:** Consistent "Mark {Category} Clusters" pattern

### ✅ **5. Complete MarkPrefixSettings Class**
- **Before:** Referenced but not defined
- **After:** Proper class with GetDisciplinePrefix method

## Implementation Checklist

- [x] Create proper `MarkPrefixSettings.cs` class
- [ ] Create `MarkParameterCommand.cs` with corrected transaction naming
- [ ] Create `MarkParameterService.cs` with safe filtering and continuous numbering
- [ ] Update `SleevePlacementExternalEvent.cs` to use instance variables
- [ ] Update `EmergencyMainDialog.cs` to pass data via SetContext method
- [ ] Remove static properties from EmergencyMainDialog
- [ ] Test single category processing
- [ ] Test multiple category processing
- [ ] Test mark numbering continuity on re-runs
- [ ] Test error handling scenarios

## Summary

This corrected implementation addresses all the critical issues:

1. **Thread Safety** - Uses instance variables instead of static properties
2. **Data Integrity** - No dangerous fallbacks or state persistence issues
3. **Numbering Continuity** - Checks existing marks to avoid duplicates
4. **Proper Encapsulation** - Clean separation of concerns
5. **Error Handling** - Safe fallbacks and proper error logging

The architecture is now robust, thread-safe, and follows Revit API best practices.
