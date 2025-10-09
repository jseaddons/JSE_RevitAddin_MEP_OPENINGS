# MEPMARK Integration Implementation Plan

## Overview
This document outlines the implementation plan for integrating MEPMARK parameter application immediately after clustering each category in the main UI OnOK click flow.

## Current Architecture Analysis

### Existing Transaction Management Pattern
Based on analysis of `UniversalClusterCommand` and `SleevePlacementExternalEvent`:

```csharp
// Pattern: Each command manages its own transaction
public void Execute(UIApplication app)
{
    using (var tx = new Transaction(doc, $"Cluster {_targetCategory} Openings"))
    {
        tx.Start();
        // ... perform operations ...
        tx.Commit();
    }
}
```

### Current Flow in SleevePlacementExternalEvent
```csharp
foreach (var category in _selectedCategories)
{
    // Step 1: Place individual sleeves
    var placementCommand = new UniversalSleevePlacementCommand(category);
    placementCommand.Execute(app);
    
    // Step 2: Cluster sleeves
    var clusterCommand = new UniversalClusterCommand(category);
    clusterCommand.Execute(app);
    
    // ⚠️ MISSING: Step 3 - Apply MEPMARK
}
```

## Implementation Plan

### Step 1: Create MarkParameterCommand

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
                using (var tx = new Transaction(doc, $"Apply MEPMARK to {_targetCategory} Clusters"))
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

### Step 2: Create MarkParameterService

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
                
                // Apply MEPMARK to each cluster
                foreach (var cluster in clusterSleeves)
                {
                    try
                    {
                        string markValue = GenerateMarkValue(category, disciplinePrefix, processedCount + 1);
                        string fullMarkValue = $"{projectPrefix}{markValue}";
                        
                        SetMarkParameter(cluster, fullMarkValue);
                        processedCount++;
                        
                        DebugLogger.Info($"[MarkParameterService] Applied MEPMARK '{fullMarkValue}' to cluster {cluster.Id.IntegerValue}");
                    }
                    catch (Exception ex)
                    {
                        errorCount++;
                        DebugLogger.Error($"[MarkParameterService] Error applying MEPMARK to cluster {cluster.Id.IntegerValue}: {ex.Message}");
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
            
            // If no category parameter, fall back to all cluster sleeves
            // (This handles cases where MEP_Category parameter is not set)
            if (categorySleeves.Count == 0)
            {
                DebugLogger.Warning($"[MarkParameterService] No sleeves found with MEP_Category='{category}', falling back to all cluster sleeves");
                return allClusterSleeves;
            }
            
            return categorySleeves;
        }
        
        /// <summary>
        /// Generate mark value based on category and index
        /// </summary>
        private string GenerateMarkValue(string category, string disciplinePrefix, int index)
        {
            // Map category to discipline prefix if not provided
            string effectivePrefix = !string.IsNullOrEmpty(disciplinePrefix) ? disciplinePrefix : GetDefaultDisciplinePrefix(category);
            
            return $"{effectivePrefix}{index:000}";
        }
        
        /// <summary>
        /// Get default discipline prefix for category
        /// </summary>
        private string GetDefaultDisciplinePrefix(string category)
        {
            return category switch
            {
                "Ducts" => "DCT",
                "Pipes" => "PLU",
                "Cable Trays" => "ELE",
                "Duct Accessories" => "DAM",
                _ => "OPN"
            };
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

### Step 3: Integrate into SleevePlacementExternalEvent

**File:** `Services/SleevePlacementExternalEvent.cs`

**Changes to existing code:**

```csharp
// In Execute method, after clusterCommand.Execute(app):
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
            
            // ⚠️ NEW: Step 3: Apply MEPMARK to clusters
            DebugLogger.Info($"[SleevePlacementExternalEvent] Applying MEPMARK to {category} clusters...");
            var markPrefixes = ReadMarkPrefixesFromUI();
            string projectPrefix = markPrefixes.ProjectPrefix;
            string disciplinePrefix = markPrefixes.GetDisciplinePrefix(category);
            
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
```

### Step 4: Add ReadMarkPrefixesFromUI method to SleevePlacementExternalEvent

**File:** `Services/SleevePlacementExternalEvent.cs`

**Add this method:**

```csharp
/// <summary>
/// Read mark prefix settings from UI (called from ExternalEvent context)
/// ⚠️ NOTE: This needs to access UI state from ExternalEvent context
/// </summary>
private MarkPrefixSettings ReadMarkPrefixesFromUI()
{
    var markPrefixes = new MarkPrefixSettings();
    
    try
    {
        // ⚠️ CRITICAL: We need to access UI state from ExternalEvent context
        // This requires passing UI references or using a different approach
        
        // Option A: Pass UI references to ExternalEvent
        // Option B: Store prefix values in static properties
        // Option C: Read from XML files (persistent storage)
        
        // For now, using Option B (static properties) for simplicity
        // TODO: Implement proper UI state access from ExternalEvent context
        
        markPrefixes.ProjectPrefix = GetProjectPrefixFromStatic();
        markPrefixes.DuctPrefix = GetDisciplinePrefixFromStatic("Ducts");
        markPrefixes.PipePrefix = GetDisciplinePrefixFromStatic("Pipes");
        markPrefixes.CableTrayPrefix = GetDisciplinePrefixFromStatic("Cable Trays");
        markPrefixes.DamperPrefix = GetDisciplinePrefixFromStatic("Duct Accessories");
        
        DebugLogger.Info($"[SleevePlacementExternalEvent] Read mark prefixes: Project='{markPrefixes.ProjectPrefix}', Ducts='{markPrefixes.DuctPrefix}', Pipes='{markPrefixes.PipePrefix}'");
    }
    catch (Exception ex)
    {
        DebugLogger.Error($"[SleevePlacementExternalEvent] Error reading mark prefixes: {ex.Message}");
        // Return defaults on error
        markPrefixes.ProjectPrefix = "SLEEVE_";
        markPrefixes.DuctPrefix = "DCT";
        markPrefixes.PipePrefix = "PLU";
        markPrefixes.CableTrayPrefix = "ELE";
        markPrefixes.DamperPrefix = "DAM";
    }
    
    return markPrefixes;
}

// ⚠️ TEMPORARY: Static properties for UI state access
// TODO: Replace with proper UI state access mechanism
private static string GetProjectPrefixFromStatic()
{
    // This needs to be set by the UI before calling ExternalEvent
    return EmergencyMainDialog.LastProjectPrefix ?? "SLEEVE_";
}

private static string GetDisciplinePrefixFromStatic(string category)
{
    // This needs to be set by the UI before calling ExternalEvent
    return EmergencyMainDialog.LastDisciplinePrefixes?.GetValueOrDefault(category) ?? "OPN";
}
```

### Step 5: Update EmergencyMainDialog to pass prefix values

**File:** `Views/EmergencyMainDialog.cs`

**Changes to OnOkClick method:**

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
        
        // ⚠️ NEW: Read and store mark prefixes before starting ExternalEvent
        var markPrefixes = ReadMarkPrefixesFromUI();
        
        // Store prefix values for ExternalEvent access
        EmergencyMainDialog.LastProjectPrefix = markPrefixes.ProjectPrefix;
        EmergencyMainDialog.LastDisciplinePrefixes = new Dictionary<string, string>
        {
            { "Ducts", markPrefixes.DuctPrefix },
            { "Pipes", markPrefixes.PipePrefix },
            { "Cable Trays", markPrefixes.CableTrayPrefix },
            { "Duct Accessories", markPrefixes.DamperPrefix }
        };
        
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
        
        // Set selected categories in external event
        _sleevePlacementHandler.SetSelectedCategories(selectedCategories);
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

// ⚠️ NEW: Static properties for ExternalEvent access
public static string LastProjectPrefix { get; set; } = "SLEEVE_";
public static Dictionary<string, string> LastDisciplinePrefixes { get; set; } = new Dictionary<string, string>();
```

## Transaction Management Strategy

### Pattern: Each Command Manages Its Own Transaction

```csharp
// ✅ CORRECT: Each command starts and commits its own transaction
UniversalSleevePlacementCommand.Execute() // Transaction 1: Place sleeves
UniversalClusterCommand.Execute()         // Transaction 2: Cluster sleeves  
MarkParameterCommand.Execute()           // Transaction 3: Apply MEPMARK
```

### Why This Approach:

✅ **Isolation** - Each operation is independent  
✅ **Rollback safety** - If one operation fails, others are not affected  
✅ **Debugging** - Easy to identify which operation failed  
✅ **Existing pattern** - Follows current architecture  
✅ **Revit best practices** - Short transactions are preferred  

## Error Handling Strategy

### Command Level:
- Each command catches exceptions and logs them
- Commands re-throw exceptions to let orchestrator handle them

### Orchestrator Level:
- SleevePlacementExternalEvent catches exceptions and shows user-friendly messages
- Continues processing other categories even if one fails

### UI Level:
- OnOkClick validates input and shows validation errors
- Stores prefix values before starting ExternalEvent

## Testing Strategy

### Test Scenarios:

1. **Single Category Processing:**
   - Select "Ducts" only
   - Verify: MEPMARK applied with correct discipline prefix

2. **Multiple Category Processing:**
   - Select "Ducts", "Pipes", "Cable Trays"
   - Verify: Each category gets correct discipline prefix

3. **Error Handling:**
   - Test with invalid prefix values
   - Test with no cluster sleeves found
   - Test with transaction failures

4. **UI Integration:**
   - Test prefix values from UI textboxes
   - Test discipline prefix updates from dropdown

## Implementation Checklist

- [ ] Create `MarkParameterCommand.cs`
- [ ] Create `MarkParameterService.cs`
- [ ] Update `SleevePlacementExternalEvent.cs` to call MarkParameterCommand
- [ ] Add static properties to `EmergencyMainDialog.cs` for prefix storage
- [ ] Update `OnOkClick` method to store prefix values
- [ ] Test single category processing
- [ ] Test multiple category processing
- [ ] Test error handling scenarios
- [ ] Verify transaction management
- [ ] Update documentation

## Future Enhancements

1. **Parameter Transfer Service Integration:**
   - Store cluster metadata for later parameter transfer
   - Support for recursive clustering between categories

2. **Advanced Mark Generation:**
   - Support for custom mark patterns
   - Integration with project numbering systems

3. **UI Improvements:**
   - Real-time preview of mark values
   - Validation of prefix formats

## Conclusion

This implementation plan provides a robust, transaction-safe approach to applying MEPMARK parameters immediately after clustering each category. The design follows existing architectural patterns and provides clear separation of concerns while maintaining error handling and user feedback capabilities.
