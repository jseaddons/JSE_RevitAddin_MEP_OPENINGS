using System.Collections.ObjectModel;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.ViewModels
{
    public class FamilySelectionViewModel
    {
    public ObservableCollection<FamilySymbol> AvailableFamilies { get; } = new ObservableCollection<FamilySymbol>();
    public FamilySymbol? SelectedFamily { get; set; } = null;

        public FamilySelectionViewModel(Document doc)
        {
            // Find all Generic Model family symbols (you can filter further if needed)
            var symbols = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .Cast<FamilySymbol>();

            foreach (var s in symbols)
                AvailableFamilies.Add(s);
        }
    }
}
