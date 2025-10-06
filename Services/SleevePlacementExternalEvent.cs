using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// External Event Handler - JUST a transaction context bridge
    /// Only provides proper Revit API context and tells orchestrator which categories to process
    /// Orchestrator routes to individual commands, commands get their own data
    /// </summary>
    public class SleevePlacementExternalEvent : IExternalEventHandler
    {
        private List<string> _selectedCategories;
        private Document _document;
        private UIDocument _uiDocument;

        public void Execute(UIApplication app)
        {
            var sw = System.Diagnostics.Stopwatch.StartNew();
            try
            {
                _uiDocument = app.ActiveUIDocument;
                _document = _uiDocument.Document;
                
                // 🎯 IMMEDIATE FEEDBACK: Show processing started
                TaskDialog.Show("Processing Started", 
                    $"Starting sleeve placement for {_selectedCategories.Count} categories:\n{string.Join(", ", _selectedCategories)}\n\nThis may take several minutes...");
                
                using (var orch = new OpeningCommandOrchestrator(_document, _uiDocument))
                {
                    var cts = new System.Threading.CancellationTokenSource(30000); // 30 seconds instead of 5
                    orch.ExecuteCategories(_selectedCategories, cts.Token);
                }
                
                // 🎯 COMPLETION FEEDBACK: Show success
                sw.Stop();
                TaskDialog.Show("Processing Complete", 
                    $"Sleeve placement completed successfully!\n\nCategories processed: {string.Join(", ", _selectedCategories)}\nProcessing time: {sw.Elapsed.TotalSeconds:F1} seconds\n\nCheck log files for detailed results.");
            }
            catch (System.OperationCanceledException)
            {
                sw.Stop();
                var message = $"Processing ABORTED after {sw.Elapsed.TotalSeconds:F1} seconds.\nCategories: {string.Join(", ", _selectedCategories)}\n\nThis may be due to timeout or cancellation.";
                System.IO.File.WriteAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\spin.log",
                    $"ABORTED after {sw.Elapsed.TotalSeconds:F1} s\r\n" +
                    $"{string.Join(", ", _selectedCategories)}");
                
                // 🎯 ERROR FEEDBACK: Show timeout/cancellation
                TaskDialog.Show("Processing Aborted", message);
            }
            catch (System.Exception ex)
            {
                sw.Stop();
                var message = $"Processing FAILED after {sw.Elapsed.TotalSeconds:F1} seconds.\nError: {ex.Message}\nCategories: {string.Join(", ", _selectedCategories)}";
                System.IO.File.WriteAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\spin.log",
                    $"CRASH after {sw.Elapsed.TotalSeconds:F1} s\r\n{ex}");
                
                // 🎯 ERROR FEEDBACK: Show failure
                TaskDialog.Show("Processing Failed", message);
            }
            finally
            {
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\spin.log",
                    $"\r\nEND at {System.DateTime.Now:HH:mm:ss}");
            }
        }

        public string GetName()
        {
            return "Sleeve Placement Handler";
        }

        public void SetSelectedCategories(List<string> selectedCategories)
        {
            _selectedCategories = selectedCategories;
            DebugLogger.Info($"[SleevePlacementExternalEvent] Set categories for processing: {string.Join(", ", selectedCategories)}");
        }

    }
}

