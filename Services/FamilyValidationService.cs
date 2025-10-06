using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Service to validate that required families are loaded in the Revit project
    /// </summary>
    public class FamilyValidationService
    {
        private readonly Document _document;
        private readonly Action<string> _logAction;

        public FamilyValidationService(Document document, Action<string> logAction)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _logAction = logAction ?? throw new ArgumentNullException(nameof(logAction));
        }

        /// <summary>
        /// Validates that all required families for opening creation are loaded
        /// </summary>
        public void ValidateRequiredFamilies()
        {
            _logAction("=== FAMILY VALIDATION STARTED ===");
            
            var missingFamilies = new List<string>();
            
            // Check for duct sleeve families
            if (!ValidateDuctSleeveFamilies())
            {
                missingFamilies.Add("Duct Sleeve Families (OpeningOnWall, OpeningOnSlab with DS# symbols)");
            }
            
            // Check for pipe sleeve families
            if (!ValidatePipeSleeveFamilies())
            {
                missingFamilies.Add("Pipe Sleeve Families (PipeOpeningOnWall, PipeOpeningOnSlab)");
            }
            
            // Check for cable tray sleeve families
            if (!ValidateCableTraySleeveFamilies())
            {
                missingFamilies.Add("Cable Tray Sleeve Families (CableTrayOpeningOnWall, CableTrayOpeningOnSlab)");
            }
            
            // Check for fire damper families
            if (!ValidateFireDamperFamilies())
            {
                missingFamilies.Add("Fire Damper Families (FireDamper, MSFD)");
            }
            
            if (missingFamilies.Count > 0)
            {
                _logAction($"Missing families detected: {missingFamilies.Count}");
                ShowMissingFamiliesDialog(missingFamilies);
            }
            else
            {
                _logAction("All required families are loaded ✓");
            }
            
            _logAction("=== FAMILY VALIDATION COMPLETED ===");
        }

        private bool ValidateDuctSleeveFamilies()
        {
            _logAction("Checking duct sleeve families...");
            
            var wallSymbol = new FilteredElementCollector(_document)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault(sym => sym.Family.Name.Contains("OpeningOnWall")
                    && sym.Name.Replace(" ", "").StartsWith("DS#", StringComparison.OrdinalIgnoreCase));
            
            var slabSymbol = new FilteredElementCollector(_document)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault(sym => sym.Family.Name.Contains("OpeningOnSlab")
                    && sym.Name.Replace(" ", "").StartsWith("DS#", StringComparison.OrdinalIgnoreCase));
            
            bool wallFound = wallSymbol != null;
            bool slabFound = slabSymbol != null;
            
            _logAction($"Duct wall symbol found: {wallFound} (ID: {wallSymbol?.Id.IntegerValue ?? 0})");
            _logAction($"Duct slab symbol found: {slabFound} (ID: {slabSymbol?.Id.IntegerValue ?? 0})");
            
            return wallFound && slabFound;
        }

        private bool ValidatePipeSleeveFamilies()
        {
            _logAction("Checking pipe sleeve families...");
            
            var wallSymbol = new FilteredElementCollector(_document)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault(sym => sym.Family.Name.Contains("PipeOpeningOnWall"));
            
            var slabSymbol = new FilteredElementCollector(_document)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault(sym => sym.Family.Name.Contains("PipeOpeningOnSlab"));
            
            bool wallFound = wallSymbol != null;
            bool slabFound = slabSymbol != null;
            
            _logAction($"Pipe wall symbol found: {wallFound} (ID: {wallSymbol?.Id.IntegerValue ?? 0})");
            _logAction($"Pipe slab symbol found: {slabFound} (ID: {slabSymbol?.Id.IntegerValue ?? 0})");
            
            return wallFound && slabFound;
        }

        private bool ValidateCableTraySleeveFamilies()
        {
            _logAction("Checking cable tray sleeve families...");
            
            var wallSymbol = new FilteredElementCollector(_document)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault(sym => string.Equals(sym.Family.Name, "CableTrayOpeningOnWall", StringComparison.OrdinalIgnoreCase));
            
            var slabSymbol = new FilteredElementCollector(_document)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault(sym => string.Equals(sym.Family.Name, "CableTrayOpeningOnSlab", StringComparison.OrdinalIgnoreCase));
            
            bool wallFound = wallSymbol != null;
            bool slabFound = slabSymbol != null;
            
            _logAction($"Cable tray wall symbol found: {wallFound} (ID: {wallSymbol?.Id.IntegerValue ?? 0})");
            _logAction($"Cable tray slab symbol found: {slabFound} (ID: {slabSymbol?.Id.IntegerValue ?? 0})");
            
            return wallFound && slabFound;
        }

        private bool ValidateFireDamperFamilies()
        {
            _logAction("Checking fire damper families...");
            
            var fireDamperSymbol = new FilteredElementCollector(_document)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault(sym => sym.Family.Name.Contains("FireDamper") || 
                                      sym.Family.Name.Contains("MSFD") ||
                                      sym.Name.Contains("STANDARD", StringComparison.OrdinalIgnoreCase));
            
            bool found = fireDamperSymbol != null;
            
            _logAction($"Fire damper symbol found: {found} (ID: {fireDamperSymbol?.Id.IntegerValue ?? 0})");
            
            return found;
        }

        private void ShowMissingFamiliesDialog(List<string> missingFamilies)
        {
            string message = "The following required families are missing from your project:\n\n";
            message += string.Join("\n• ", missingFamilies);
            message += "\n\nPlease load these families using:\n";
            message += "Insert > Load Family\n\n";
            message += "The families should be in the Generic Model category.\n\n";
            message += "Without these families, opening creation will fail.";
            
            _logAction($"Showing missing families dialog: {missingFamilies.Count} families missing");
            
            TaskDialog.Show("Missing Required Families", message);
        }
    }
}
