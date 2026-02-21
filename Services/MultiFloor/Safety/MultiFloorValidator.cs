using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.MultiFloor.Safety
{
    /// <summary>
    /// Validation for multi-floor operations.
    /// Prevents processing invalid data that could cause crashes.
    /// </summary>
    public class MultiFloorValidator
    {
        private readonly Action<string> _logger;
        
        public MultiFloorValidator(Action<string> logger = null)
        {
            _logger = logger ?? (msg => DebugLogger.Info(msg));
        }
        
        /// <summary>
        /// Validates a level before processing
        /// </summary>
        public ValidationResult ValidateLevel(Level level, Document doc)
        {
            var errors = new List<string>();
            var warnings = new List<string>();
            
            if (level == null)
            {
                errors.Add("Level is null");
                return new ValidationResult { IsValid = false, Errors = errors };
            }
            
            // Check level ID
            if (level.Id == null || level.Id.GetIntegerValue() <= 0)
            {
                errors.Add($"Level '{level.Name}' has invalid ID");
            }
            
            // Check level name
            if (string.IsNullOrWhiteSpace(level.Name))
            {
                errors.Add("Level has no name");
            }
            
            // Check elevation
            try
            {
                var elevation = level.ProjectElevation;
                if (double.IsNaN(elevation) || double.IsInfinity(elevation))
                {
                    errors.Add($"Level '{level.Name}' has invalid elevation: {elevation}");
                }
                else if (Math.Abs(elevation) > 100000) // 100,000 feet is unrealistic
                {
                    warnings.Add($"Level '{level.Name}' has extreme elevation: {elevation:F2} ft");
                }
            }
            catch (Exception ex)
            {
                errors.Add($"Level '{level.Name}' elevation access failed: {ex.Message}");
            }
            
            // Check level is not a view-specific element
            try
            {
                var isViewSpecific = level.ViewSpecific;
                // Levels should not be view-specific
            }
            catch
            {
                warnings.Add($"Level '{level.Name}' view-specific check failed");
            }
            
            // Check document association
            if (doc == null)
            {
                errors.Add("Document is null");
            }
            else if (doc.IsReadOnly)
            {
                errors.Add("Document is read-only");
            }
            else if (!doc.IsValidObject)
            {
                errors.Add("Document is no longer valid");
            }
            
            return new ValidationResult 
            { 
                IsValid = errors.Count == 0, 
                Errors = errors,
                Warnings = warnings
            };
        }
        
        /// <summary>
        /// Validates a list of clash zones before processing
        /// </summary>
        public ValidationResult ValidateClashZones(List<ClashZone> clashes, Document doc)
        {
            var errors = new List<string>();
            var warnings = new List<string>();
            var validClashes = new List<ClashZone>();
            
            if (clashes == null)
            {
                errors.Add("Clash list is null");
                return new ValidationResult { IsValid = false, Errors = errors };
            }
            
            if (clashes.Count == 0)
            {
                // Empty is valid, just nothing to do
                return new ValidationResult { IsValid = true, ValidClashes = validClashes };
            }
            
            foreach (var clash in clashes)
            {
                if (clash == null)
                {
                    warnings.Add("Null clash zone in list, skipping");
                    continue;
                }
                
                // Validate clash zone ID
                if (clash.Id == Guid.Empty)
                {
                    warnings.Add("Clash zone has no ID, skipping");
                    continue;
                }
                var clashIdStr = clash.Id.ToString();
                
                // Validate MEP element exists
                if (clash.MepElementId == null || clash.MepElementId == ElementId.InvalidElementId)
                {
                    warnings.Add($"Clash {clashIdStr} has no MEP element, skipping");
                    continue;
                }
                
                // Check MEP element still exists in document
                bool elementExists = true;
                try
                {
                    var mepElement = doc.GetElement(clash.MepElementId);
                    if (mepElement == null || !mepElement.IsValidObject)
                    {
                        elementExists = false;
                    }
                }
                catch
                {
                    elementExists = false;
                }
                
                if (!elementExists)
                {
                    warnings.Add($"Clash {clashIdStr} has deleted MEP element ({clash.MepElementId}), skipping");
                    continue;
                }
                
                // Validate placement point
                if (clash.SleevePlacementPoint == null)
                {
                    warnings.Add($"Clash {clashIdStr} has no placement point, skipping");
                    continue;
                }
                
                // Check for extreme coordinates (possible corruption)
                var pt = clash.SleevePlacementPoint;
                if (Math.Abs(pt.X) > 1000000 || Math.Abs(pt.Y) > 1000000 || Math.Abs(pt.Z) > 1000000)
                {
                    warnings.Add($"Clash {clashIdStr} has extreme coordinates ({pt.X:F0}, {pt.Y:F0}, {pt.Z:F0}), possible corruption");
                    // Still include it but warn
                }
                
                validClashes.Add(clash);
            }
            
            _logger?.Invoke($"[Validator] Validated {clashes.Count} clashes: {validClashes.Count} valid, {clashes.Count - validClashes.Count} filtered out");
            
            return new ValidationResult 
            { 
                IsValid = errors.Count == 0, 
                Errors = errors,
                Warnings = warnings,
                ValidClashes = validClashes
            };
        }
        
        /// <summary>
        /// Pre-flight validation before starting multi-floor operation
        /// </summary>
        public ValidationResult ValidateBeforeProcessing(List<Level> levels, Document doc)
        {
            var errors = new List<string>();
            var warnings = new List<string>();
            
            if (levels == null || levels.Count == 0)
            {
                errors.Add("No levels provided for processing");
                return new ValidationResult { IsValid = false, Errors = errors };
            }
            
            // Check document
            if (doc == null)
            {
                errors.Add("Document is null");
                return new ValidationResult { IsValid = false, Errors = errors };
            }
            
            if (doc.IsReadOnly)
            {
                errors.Add("Document is read-only");
            }
            
            if (!doc.IsValidObject)
            {
                errors.Add("Document is no longer valid");
            }
            
            // Check for too many levels (memory risk)
            if (levels.Count > 100)
            {
                warnings.Add($"Processing {levels.Count} levels may cause memory issues. Consider processing in smaller batches.");
            }
            
            // Check for duplicate level names
            var duplicateNames = levels
                .GroupBy(l => l.Name)
                .Where(g => g.Count() > 1)
                .Select(g => g.Key)
                .ToList();
                
            if (duplicateNames.Any())
            {
                warnings.Add($"Duplicate level names found: {string.Join(", ", duplicateNames)}. Processing may be ambiguous.");
            }
            
            // Validate each level
            int invalidLevelCount = 0;
            foreach (var level in levels)
            {
                var levelValidation = ValidateLevel(level, doc);
                if (!levelValidation.IsValid)
                {
                    invalidLevelCount++;
                    errors.AddRange(levelValidation.Errors);
                }
                warnings.AddRange(levelValidation.Warnings);
            }
            
            if (invalidLevelCount > 0)
            {
                errors.Add($"{invalidLevelCount} of {levels.Count} levels failed validation");
            }
            
            return new ValidationResult 
            { 
                IsValid = errors.Count == 0, 
                Errors = errors,
                Warnings = warnings
            };
        }
    }
    
    /// <summary>
    /// Result of validation
    /// </summary>
    public class ValidationResult
    {
        public bool IsValid { get; set; }
        public List<string> Errors { get; set; } = new List<string>();
        public List<string> Warnings { get; set; } = new List<string>();
        public List<ClashZone> ValidClashes { get; set; } = new List<ClashZone>();
        
        public string GetSummary()
        {
            if (IsValid && Warnings.Count == 0)
                return "Validation passed";
                
            var parts = new List<string>();
            if (!IsValid)
                parts.Add($"{Errors.Count} errors");
            if (Warnings.Count > 0)
                parts.Add($"{Warnings.Count} warnings");
                
            return string.Join(", ", parts);
        }
    }
}
