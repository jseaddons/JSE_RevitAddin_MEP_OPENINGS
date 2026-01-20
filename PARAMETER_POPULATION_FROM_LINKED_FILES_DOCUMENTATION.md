# Parameter Population from Linked Files Documentation

## Overview
This document explains how the parameter population system works to extract MEP parameters from linked reference files based on selected categories.

## Problem Solved
Previously, the parameter population was failing because it was trying to extract parameters from the current document instead of the linked reference files where the MEP elements actually exist.

## Architecture

### Key Components

1. **ParameterExtractionService** - Core service for parameter extraction
2. **LinkedFileService** - Service for managing linked files
3. **EmergencyMainDialog** - Main UI that orchestrates parameter population

### Flow Diagram

```
User Clicks Refresh
    ↓
EmergencyMainDialog.PopulateParameterDropdowns()
    ↓
ParameterExtractionService.PopulateCategorySpecificParameters()
    ↓
LinkedFileService.GetLinkedFiles() → Get all linked files
    ↓
For each selected MEP category:
    ↓
GetParametersForSpecificCategoryFromLinkedFiles()
    ↓
For each linked file:
    ↓
GetParametersForMepCategories(linkedDoc, category)
    ↓
UpdateSingleTabParameters() → Populate UI dropdowns
```

## Implementation Details

### 1. Main UI Method (`EmergencyMainDialog.cs`)

```csharp
private void PopulateParameterDropdowns()
{
    try
    {
        DebugLogger.Info("[PARAMETER_SERVICE] Starting category-specific parameter population via service");
        
        // Get ALL selected MEP categories
        var selectedCategories = GetSelectedMepCategories();
        DebugLogger.Info($"[PARAMETER_SERVICE] Selected MEP categories: {string.Join(", ", selectedCategories)}");
        
        // Call service to handle category-specific parameter population
        var parameterService = new Services.ParameterExtractionService();
        parameterService.PopulateCategorySpecificParameters(_serviceParameterTabs, selectedCategories, _uiDocument?.Document);
        
        DebugLogger.Info("[PARAMETER_SERVICE] Category-specific parameter population completed via service");
    }
    catch (Exception ex)
    {
        DebugLogger.Error($"[PARAMETER_SERVICE] Error in PopulateParameterDropdowns: {ex.Message}");
    }
}
```

### 2. Service Method (`ParameterExtractionService.cs`)

```csharp
public void PopulateCategorySpecificParameters(TabControl serviceParameterTabs, List<string> selectedCategories, Document document)
{
    try
    {
        // Get opening parameters (global - same for all tabs)
        var openingParameters = GetCurrentOpeningParameters(document);
        
        // Get linked files for parameter extraction
        var linkedFileService = new Services.LinkedFileService();
        var linkedFiles = linkedFileService.GetLinkedFiles(document);
        
        // Update each tab that matches selected categories
        foreach (var selectedCategory in selectedCategories)
        {
            var matchingTab = FindTabByCategory(serviceParameterTabs, selectedCategory);
            if (matchingTab != null)
            {
                var categorySpecificParameters = GetParametersForSpecificCategoryFromLinkedFiles(selectedCategory, linkedFiles);
                UpdateSingleTabParameters(matchingTab, categorySpecificParameters, openingParameters);
            }
        }
    }
    catch (Exception ex)
    {
        // Error handling
    }
}
```

### 3. Linked File Parameter Extraction

```csharp
private List<string> GetParametersForSpecificCategoryFromLinkedFiles(string categoryName, List<Services.LinkedFileInfo> linkedFiles)
{
    try
    {
        var categoryParameters = new List<string>();
        
        if (linkedFiles == null || linkedFiles.Count == 0)
        {
            return categoryParameters;
        }
        
        if (Enum.TryParse<MepCategory>(categoryName, out var category))
        {
            // Get parameters from all linked files
            foreach (var linkedFile in linkedFiles)
            {
                try
                {
                    var linkedDoc = linkedFile.LinkInstance?.GetLinkDocument();
                    if (linkedDoc != null)
                    {
                        var parameters = GetParametersForMepCategories(linkedDoc, new List<MepCategory> { category });
                        var parameterNames = parameters.Select(p => p.Name).ToList();
                        
                        // Add unique parameters
                        foreach (var paramName in parameterNames)
                        {
                            if (!categoryParameters.Contains(paramName))
                            {
                                categoryParameters.Add(paramName);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Log error but continue with other files
                }
            }
        }
        
        return categoryParameters;
    }
    catch (Exception ex)
    {
        return new List<string>();
    }
}
```

### 4. New Row Parameter Population (`EmergencyMainDialog.cs`)

```csharp
private List<string> GetCurrentMepParameters()
{
    try
    {
        var mepParameters = new List<string>();
        
        // Get selected MEP categories
        var selectedCategories = GetSelectedMepCategories();
        if (selectedCategories.Count == 0)
        {
            return mepParameters;
        }
        
        // Get current document
        var document = _uiDocument?.Document;
        if (document == null)
        {
            return mepParameters;
        }
        
        // Convert string categories to enum
        var mepCategories = new List<Services.MepCategory>();
        foreach (var categoryName in selectedCategories)
        {
            if (Enum.TryParse<Services.MepCategory>(categoryName, out var category))
            {
                mepCategories.Add(category);
            }
        }
        
        // Use ParameterExtractionService to get parameters from linked files
        var parameterService = new Services.ParameterExtractionService();
        var linkedFileService = new Services.LinkedFileService();
        var linkedFiles = linkedFileService.GetLinkedFiles(document);
        
        if (linkedFiles.Count > 0)
        {
            // Get parameters from linked files
            foreach (var linkedFile in linkedFiles)
            {
                try
                {
                    var linkedDoc = linkedFile.LinkInstance?.GetLinkDocument();
                    if (linkedDoc != null)
                    {
                        var parameters = parameterService.GetParametersForMepCategories(linkedDoc, mepCategories);
                        var parameterNames = parameters.Select(p => p.Name).ToList();
                        
                        // Add unique parameters
                        foreach (var paramName in parameterNames)
                        {
                            if (!mepParameters.Contains(paramName))
                            {
                                mepParameters.Add(paramName);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    DebugLogger.Warning($"[PARAMETER_CURRENT] Error getting parameters from linked file '{linkedFile.FileName}': {ex.Message}");
                }
            }
        }
        else
        {
            // Fallback to current document if no linked files
            var parameters = parameterService.GetParametersForMepCategories(document, mepCategories);
            mepParameters = parameters.Select(p => p.Name).ToList();
        }
        
        return mepParameters;
    }
    catch (Exception ex)
    {
        DebugLogger.Error($"[PARAMETER_CURRENT] Error getting current MEP parameters: {ex.Message}");
        return new List<string>();
    }
}
```

## Key Technical Points

### 1. Linked File Access
- **Correct Method**: `linkedFile.LinkInstance?.GetLinkDocument()`
- **Incorrect Method**: `linkedFile.Document` (doesn't exist)

### 2. Parameter Deduplication
- Parameters from multiple linked files are combined
- Duplicate parameter names are filtered out using `Contains()` check

### 3. Error Handling
- Each linked file is processed independently
- If one linked file fails, others continue processing
- Comprehensive logging for debugging

### 4. Fallback Strategy
- If no linked files are found, falls back to current document
- Ensures system doesn't break if linked files are unavailable

## Usage Instructions

### For Developers

1. **Adding New MEP Categories**:
   - Add to `MepCategory` enum in `LinkedFileDetectionService.cs`
   - Update parameter extraction logic if needed

2. **Modifying Parameter Extraction**:
   - Modify `GetParametersForMepCategories()` in `ParameterExtractionService.cs`
   - This affects both refresh and new row population

3. **Debugging Parameter Issues**:
   - Check `logger_debug.txt` for service method logs
   - Check main UI logs for overall flow
   - Verify linked files are being detected correctly

### For Users

1. **Select MEP Categories**: Check the desired MEP categories (Ducts, Pipes, Cable Trays, etc.)
2. **Click Refresh**: This populates parameters from linked files
3. **Add New Rows**: New parameter rows automatically get populated with parameters from linked files
4. **Verify Results**: Check that parameters appear in the dropdowns

## Troubleshooting

### Common Issues

1. **No Parameters Found**:
   - Check if linked files are loaded
   - Verify MEP categories are selected
   - Check if linked files contain elements of selected categories

2. **Parameters Not Updating**:
   - Restart Revit (DLL caching issue)
   - Check logs for errors
   - Verify linked file access permissions

3. **Wrong Parameters**:
   - Check if correct MEP categories are selected
   - Verify linked files contain expected elements
   - Check parameter filtering logic

### Debug Logs

- **Service Logs**: `Log/logger_debug.txt`
- **Main UI Logs**: `Log/MainUi_*.log`
- **Refresh Logs**: `Log/Refresh_*.log`

Look for these log patterns:
- `[PARAMETER_SERVICE]` - Service method execution
- `[PARAMETER_CATEGORY]` - Category-specific parameter extraction
- `[PARAMETER_CURRENT]` - New row parameter population

## Future Enhancements

1. **Caching**: Cache parameters per linked file to improve performance
2. **Filtering**: Add more sophisticated parameter filtering options
3. **Validation**: Add parameter validation before populating UI
4. **Performance**: Optimize for projects with many linked files

## Related Files

- `Services/ParameterExtractionService.cs` - Core parameter extraction logic
- `Services/LinkedFileService.cs` - Linked file management
- `Views/EmergencyMainDialog.cs` - Main UI orchestration
- `Services/LinkedFileDetectionService.cs` - MEP category definitions

---

**Last Updated**: September 23, 2025
**Version**: 1.0
**Status**: Working ✅
