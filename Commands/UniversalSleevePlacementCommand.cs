using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Strategies;
using JSE_RevitAddin_MEP_OPENINGS.Utils;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    /// <summary>
    /// Universal sleeve placement command - handles ALL MEP categories using strategy pattern
    /// Based on proven DuctSleevePlacementCommand pattern
    /// 
    /// Supports: Ducts, Pipes, Cable Trays, Duct Accessories (Dampers)
    /// Uses 4 universal families: OpeningOnWall-Rectangular, OpeningOnWall-Circular, 
    ///                           OpeningOnSlab-Rectangular, OpeningOnSlab-Circular
    /// </summary>
    public class UniversalSleevePlacementCommand : ICommand
    {
        private readonly Document _doc;
        private readonly List<ClashZone> _clashZones;
        private readonly string _category;
        private readonly ISleevePlacementStrategy _strategy;
        private readonly string _logPrefix;
        private OpeningConditions _conditions;

        public UniversalSleevePlacementCommand(Document doc, List<ClashZone> clashZones, string category)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _clashZones = clashZones ?? throw new ArgumentNullException(nameof(clashZones));
            _category = category ?? throw new ArgumentNullException(nameof(category));
            _logPrefix = $"[UniversalSleeveCommand-{category}]";
            
            // Select strategy based on category
            _strategy = CreateStrategy(category);
            
            // Load conditions from XML
            LoadConditionsFromXml();
        }

        public void Execute(UIApplication app)
        {
            try
            {
                DebugLogger.Info($"{_logPrefix} Starting sleeve placement for {_clashZones.Count} clash zones");
                
                // ---- 1. VALIDATION: Check document state (NO transaction) ----
                if (!ValidateDocument())
                {
                    DebugLogger.Error($"{_logPrefix} Document validation failed - cannot place sleeves");
                    return;
                }

                // ---- 2. VALIDATION: Family validation removed per user request ----
                // Old families no longer needed - using universal opening families only
                
                // ---- 3. SINGLE TRANSACTION: All sleeve placement ----
                // NO MEP element collection needed - all data is in ClashZone from refresh!
                DebugLogger.Info($"{_logPrefix} Starting placement for {_clashZones.Count} clash zones (zero linked file access)");
                using (var t = new Transaction(_doc, $"Place {_category} Sleeves"))
                {
                    if (t.Start() == TransactionStatus.Started)
                    {
                        // Set failure handler to auto-resolve warnings
                        var options = t.GetFailureHandlingOptions();
                        options.SetFailuresPreprocessor(new UniversalWarningSwallower());
                        t.SetFailureHandlingOptions(options);
                        DebugLogger.Info($"{_logPrefix} Transaction started with UniversalWarningSwallower enabled");
                        
                        // Place all sleeves in single transaction (zero linked file access!)
                        var placerService = new UniversalSleevePlacerService(_doc, _conditions, _strategy);
                        var result = placerService.PlaceAllSleevesInTransaction(_clashZones);
                        
                        // Commit and check status
                        var status = t.Commit();
                        if (status == TransactionStatus.Committed)
                        {
                            DebugLogger.Info($"{_logPrefix} ✓ Transaction committed successfully - Placed: {result.PlacedCount}, Skipped: {result.SkippedCount}");
                            
                            // Show success feedback
                            string message = result.PlacedCount > 0
                                ? $"✓ Successfully placed {result.PlacedCount} {_category} sleeve(s)\n✗ Skipped {result.SkippedCount} (already resolved)"
                                : $"No {_category} sleeves placed\n✗ All {result.SkippedCount} were already resolved";
                            
                            MessageBox.Show(message, $"{_category} Sleeve Placement Complete", 
                                MessageBoxButtons.OK, 
                                MessageBoxIcon.Information);
                        }
                        else
                        {
                            DebugLogger.Error($"{_logPrefix} Transaction failed to commit: {status}");
                            DebugLogger.Error($"{_logPrefix} This might be due to duplicate suppression or existing sleeves");
                            
                            MessageBox.Show($"Failed to place {_category} sleeves.\nStatus: {status}\nCheck log for details.", 
                                "Placement Failed", 
                                MessageBoxButtons.OK, 
                                MessageBoxIcon.Warning);
                        }
                    }
                    else
                    {
                        DebugLogger.Error($"{_logPrefix} Failed to start transaction");
                        MessageBox.Show($"Failed to start transaction for {_category} sleeves.", 
                            "Transaction Error", 
                            MessageBoxButtons.OK, 
                            MessageBoxIcon.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"{_logPrefix} Exception: {ex.Message}");
                DebugLogger.Error($"{_logPrefix} Stack trace: {ex.StackTrace}");
                
                MessageBox.Show($"Error placing {_category} sleeves:\n\n{ex.Message}", 
                    "Sleeve Placement Error", 
                    MessageBoxButtons.OK, 
                    MessageBoxIcon.Error);
            }
        }
        
        private ISleevePlacementStrategy CreateStrategy(string category)
        {
            return category switch
            {
                "Ducts" => new DuctPlacementStrategy(),
                "Pipes" => new PipePlacementStrategy(),
                "Cable Trays" => new CableTrayPlacementStrategy(),
                "Duct Accessories" => new DamperPlacementStrategy(_doc),
                _ => throw new ArgumentException($"Unknown category: {category}")
            };
        }
        
        private bool ValidateDocument()
        {
            DebugLogger.Info($"{_logPrefix} Document validation - Path: {_doc.PathName}");
            DebugLogger.Info($"{_logPrefix} Document validation - IsModifiable: {_doc.IsModifiable}");
            DebugLogger.Info($"{_logPrefix} Document validation - IsLinked: {_doc.IsLinked}");
            DebugLogger.Info($"{_logPrefix} Document validation - IsWorkshared: {_doc.IsWorkshared}");
            
            // Test if we can modify
            bool canModify = false;
            try
            {
                using (var testTransaction = new Transaction(_doc, "Test Modification"))
                {
                    if (testTransaction.Start() == TransactionStatus.Started)
                    {
                        canModify = true;
                        testTransaction.RollBack();
                        DebugLogger.Info($"{_logPrefix} Document can be modified");
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"{_logPrefix} Document modification test failed: {ex.Message}");
            }
            
            return canModify;
        }
        
        private void LoadConditionsFromXml()
        {
            try
            {
                var conditionsService = new ConditionsService(msg => DebugLogger.Info(msg));
                
                // Extract filter name from category (standard naming: Ventilation for Ducts, etc.)
                // This matches the XML file naming convention: Ventilation_CONDITIONS.xml for Ducts
                string filterName = _category switch
                {
                    "Ducts" => "Ventilation",
                    "Pipes" => "Plumbing", // or "WaterSystems"
                    "Cable Trays" => "DataDevices", // or "Electrical"
                    "Duct Accessories" => "Ventilation", // Uses same as Ducts
                    _ => "Ventilation" // Default fallback
                };
                
                DebugLogger.Info($"{_logPrefix} Using filter name '{filterName}' for category '{_category}'");
                
                _conditions = conditionsService.LoadConditions(filterName);
                
                if (_conditions != null)
                {
                    DebugLogger.Info($"{_logPrefix} Loaded conditions for filter '{filterName}'");
                }
                else
                {
                    DebugLogger.Warning($"{_logPrefix} No conditions found for filter '{filterName}' - using defaults");
                    _conditions = new OpeningConditions { FilterName = filterName, Category = _category };
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"{_logPrefix} Error loading conditions: {ex.Message}");
                _conditions = new OpeningConditions { Category = _category };
            }
        }
    }

    /// <summary>
    /// Failure preprocessor to auto-dismiss warnings during universal sleeve placement
    /// </summary>
    public class UniversalWarningSwallower : IFailuresPreprocessor
    {
        public FailureProcessingResult PreprocessFailures(FailuresAccessor fa)
        {
            var failures = fa.GetFailureMessages();
            foreach (var f in failures)
            {
                var description = f.GetDescriptionText();
                if (f.GetSeverity() == FailureSeverity.Warning)
                {
                    fa.DeleteWarning(f);
                    DebugLogger.Info($"[UniversalWarningSwallower] Dismissed warning: {description}");
                }
                // Also dismiss duplicate-related errors to prevent transaction rollback
                else if (f.GetSeverity() == FailureSeverity.Error && 
                         (description.Contains("duplicate") || description.Contains("already exists")))
                {
                    fa.DeleteWarning(f);
                    DebugLogger.Info($"[UniversalWarningSwallower] Dismissed duplicate error: {description}");
                }
            }
            return FailureProcessingResult.Continue;
        }
    }
}
