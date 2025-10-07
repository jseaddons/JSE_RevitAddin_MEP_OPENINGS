using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    /// <summary>
    /// Bulletproof duct sleeve placement using industry-standard pattern
    /// </summary>
    public class DuctSleevePlacementCommand : ICommand
    {
        private readonly Document _doc;
        private readonly List<ClashZone> _clashZones;
        private readonly string _logPrefix;
        private OpeningConditions _conditions; // Loaded from CONDITIONS.xml

        public DuctSleevePlacementCommand(Document doc, List<ClashZone> clashZones)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _clashZones = clashZones ?? throw new ArgumentNullException(nameof(clashZones));
            _logPrefix = "[DuctSleevePlacementCommand]";
            
            // Load conditions from XML (filter name extracted from clash zones)
            LoadConditionsFromXml();
        }

        public void Execute(UIApplication app)
        {
            try
            {
                DebugLogger.Info($"{_logPrefix} Starting duct sleeve placement for {_clashZones.Count} clash zones");

                // ---- 1. VALIDATION: Check document state ----
                DebugLogger.Info($"{_logPrefix} Document validation - Path: {_doc.PathName}");
                DebugLogger.Info($"{_logPrefix} Document validation - IsModifiable: {_doc.IsModifiable}");
                DebugLogger.Info($"{_logPrefix} Document validation - IsLinked: {_doc.IsLinked}");
                DebugLogger.Info($"{_logPrefix} Document validation - IsWorkshared: {_doc.IsWorkshared}");
                
                // Additional workset validation for workshared documents
                if (_doc.IsWorkshared)
                {
                    try
                    {
                        var worksetTable = _doc.GetWorksetTable();
                        var activeWorksetId = worksetTable.GetActiveWorksetId();
                        var activeWorkset = worksetTable.GetWorkset(activeWorksetId);
                        DebugLogger.Info($"{_logPrefix} Active workset: {activeWorkset.Name}, Kind: {activeWorkset.Kind}, Owner: {activeWorkset.Owner}");
                        
                        // Check if the active workset is editable
                        if (activeWorkset.Kind == WorksetKind.UserWorkset && string.IsNullOrEmpty(activeWorkset.Owner))
                        {
                            DebugLogger.Warning($"{_logPrefix} Active workset '{activeWorkset.Name}' is not owned by anyone - this might prevent modification");
                        }
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Warning($"{_logPrefix} Could not check workset details: {ex.Message}");
                    }
                }
                
                // Try to start a test transaction to see if we can actually modify the document
                bool canModify = false;
                try
                {
                    using (var testTransaction = new Transaction(_doc, "Test Modification"))
                    {
                        if (testTransaction.Start() == TransactionStatus.Started)
                        {
                            canModify = true;
                            testTransaction.RollBack(); // Rollback immediately
                            DebugLogger.Info($"{_logPrefix} Test transaction successful - document can be modified");
                        }
                        else
                        {
                            DebugLogger.Error($"{_logPrefix} Test transaction failed to start");
                        }
                    }
                }
                catch (Exception ex)
                {
                    DebugLogger.Error($"{_logPrefix} Test transaction exception: {ex.Message}");
                }
                
                if (!canModify)
                {
                    var msg = $"Cannot place sleeves: Document modification test failed. " +
                             $"Document Path: {_doc.PathName}, IsLinked: {_doc.IsLinked}, IsWorkshared: {_doc.IsWorkshared}, IsModifiable: {_doc.IsModifiable}";
                    DebugLogger.Error($"{_logPrefix} {msg}");
                    // TaskDialog removed - non-blocking UI, errors logged
                    return;
                }

                // ---- 2. READ-ONLY: Data gathering (NO transaction) ----
                // Clearance will be read from CONDITIONS.xml per-duct based on shape and insulation
                var placerService = new DuctSleevePlacerService(_doc, _conditions);
                var ductsWithClashZones = placerService.CollectDuctsWithClashZones(_clashZones);

                if (ductsWithClashZones.Count == 0)
                {
                    DebugLogger.Warning($"{_logPrefix} No ducts found for clash zones");
                    // TaskDialog removed - non-blocking UI
                    return;
                }

                DebugLogger.Info($"{_logPrefix} Found {ductsWithClashZones.Count} ducts to process");

                // ---- 3. SINGLE TRANSACTION: All sleeve placement ----
                using (var t = new Transaction(_doc, "Place Duct Sleeves"))
                {
                    if (t.Start() == TransactionStatus.Started)
                    {
                        // Set failure handler to auto-resolve warnings
                        var options = t.GetFailureHandlingOptions();
                        options.SetFailuresPreprocessor(new WarningSwallower());
                        t.SetFailureHandlingOptions(options);

                        // Get duct opening type from CONDITIONS.xml
                        string ductOpeningType = _conditions?.OpeningTypePreferences?.RoundDucts ?? "Circular";
                        DebugLogger.Info($"{_logPrefix} Using duct opening type from CONDITIONS.xml: {ductOpeningType}");

                        int placed = placerService.PlaceAllSleevesOptimized(ductsWithClashZones, t, _clashZones, ductOpeningType);

                        var status = t.Commit();
                        if (status == TransactionStatus.Committed)
                        {
                            DebugLogger.Info($"{_logPrefix} ✓ Successfully placed {placed} duct sleeves");
                            
                            // Update UI status label (non-blocking)
                            UpdateStatusLabel($"✓ Placed {placed} duct sleeves");
                            
                            // Show completion prompt (blocking is OK here - placement is done)
                            System.Windows.Forms.MessageBox.Show(
                                $"Successfully placed {placed} duct sleeve(s).",
                                "Placement Complete",
                                System.Windows.Forms.MessageBoxButtons.OK,
                                System.Windows.Forms.MessageBoxIcon.Information);
                        }
                        else
                        {
                            DebugLogger.Error($"{_logPrefix} Transaction failed to commit: {status}");
                            DebugLogger.Error($"{_logPrefix} This might be due to duplicate suppression or existing sleeves at the same location");
                            
                            // Update UI status label (non-blocking)
                            UpdateStatusLabel($"✗ Failed to place sleeves - Status: {status}");
                            
                            // Show error prompt
                            System.Windows.Forms.MessageBox.Show(
                                $"Failed to place sleeves. Status: {status}\nCheck log for details.",
                                "Placement Failed",
                                System.Windows.Forms.MessageBoxButtons.OK,
                                System.Windows.Forms.MessageBoxIcon.Warning);
                        }
                    }
                    else
                    {
                        DebugLogger.Error($"{_logPrefix} Failed to start transaction");
                        // TaskDialog removed - non-blocking UI
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"{_logPrefix} Exception during sleeve placement: {ex.Message}");
                DebugLogger.Error($"{_logPrefix} Stack trace: {ex.StackTrace}");
                // TaskDialog removed - non-blocking UI
            }
        }
        
        /// <summary>
        /// ⚠️ CRITICAL METHOD - DO NOT REMOVE ⚠️
        /// Load opening conditions from XML file
        /// Filter name is "Ventilation" for "Ventilation_ducts.xml" clash zones
        /// </summary>
        private void LoadConditionsFromXml()
        {
            try
            {
                // Try to determine filter name from clash zone file path or default to "Ventilation"
                string filterName = "Ventilation"; // Default
                
                // Try to extract filter name from first clash zone's document path or geometry hash
                if (_clashZones != null && _clashZones.Count > 0)
                {
                    var firstZone = _clashZones[0];
                    // The filter name is typically part of the XML file name that loaded these zones
                    // For now, use default "Ventilation" - can be enhanced to detect from XML path
                }
                
                var conditionsService = new ConditionsService(msg => DebugLogger.Info(msg));
                _conditions = conditionsService.LoadConditions(filterName);
                
                DebugLogger.Info($"{_logPrefix} Loaded conditions from XML for filter '{filterName}'");
                DebugLogger.Info($"{_logPrefix} Clearances - RectNormal:{_conditions.ClearanceSettings.RectangularNormal}mm, RectInsulated:{_conditions.ClearanceSettings.RectangularInsulated}mm, RoundNormal:{_conditions.ClearanceSettings.RoundNormal}mm, RoundInsulated:{_conditions.ClearanceSettings.RoundInsulated}mm");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"{_logPrefix} Error loading conditions from XML: {ex.Message}");
                // Create default conditions
                _conditions = new OpeningConditions
                {
                    FilterName = "Ventilation",
                    Category = "Ducts"
                };
            }
        }
        
        private void UpdateStatusLabel(string message)
        {
            try
            {
                var mainDialog = System.Windows.Forms.Application.OpenForms.OfType<Views.EmergencyMainDialog>().FirstOrDefault();
                if (mainDialog != null)
                {
                    mainDialog.Invoke(new Action(() =>
                    {
                        var statusLabel = mainDialog.Controls.Find("_statusLabel", true).FirstOrDefault() as System.Windows.Forms.Label;
                        if (statusLabel != null)
                        {
                            statusLabel.Text = message;
                        }
                    }));
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Warning($"[UpdateStatusLabel] Could not update status: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// Failure preprocessor to auto-dismiss warnings and duplicate-related errors during bulk operations
    /// </summary>
    public class WarningSwallower : IFailuresPreprocessor
    {
        public FailureProcessingResult PreprocessFailures(FailuresAccessor fa)
        {
            var failures = fa.GetFailureMessages();
            foreach (var f in failures)
            {
                string description = f.GetDescriptionText();
                
                // Dismiss warnings (as before)
                if (f.GetSeverity() == FailureSeverity.Warning)
                {
                    fa.DeleteWarning(f);
                    DebugLogger.Info($"[WarningSwallower] Dismissed warning: {description}");
                }
                // CRITICAL FIX: Also dismiss duplicate-related errors to prevent transaction rollback
                else if (f.GetSeverity() == FailureSeverity.Error)
                {
                    // Check if this is a duplicate-related error that should be suppressed
                    if (IsDuplicateRelatedError(description))
                    {
                        fa.DeleteWarning(f);
                        DebugLogger.Info($"[WarningSwallower] Dismissed duplicate-related error: {description}");
                    }
                    else
                    {
                        DebugLogger.Warning($"[WarningSwallower] Keeping non-duplicate error: {description}");
                    }
                }
            }
            return FailureProcessingResult.Continue;
        }
        
        /// <summary>
        /// Check if an error message is related to duplicate placement that should be suppressed
        /// </summary>
        private bool IsDuplicateRelatedError(string description)
        {
            if (string.IsNullOrEmpty(description)) return false;
            
            string lowerDesc = description.ToLower();
            
            // Common duplicate-related error patterns
            return lowerDesc.Contains("identical instances") ||
                   lowerDesc.Contains("duplicate") ||
                   lowerDesc.Contains("same place") ||
                   lowerDesc.Contains("already exists") ||
                   lowerDesc.Contains("element already") ||
                   lowerDesc.Contains("instance already") ||
                   lowerDesc.Contains("overlapping") ||
                   lowerDesc.Contains("conflict");
        }
    }
}
