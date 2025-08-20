
using System.Windows.Input;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using JSE_RevitAddin_MEP_OPENINGS.Commands;

namespace JSE_RevitAddin_MEP_OPENINGS.ViewModels
{
    public sealed class JSE_RevitAddin_MEP_OPENINGSViewModel : ObservableObject
    {
    private readonly ExternalCommandData _commandData;

    // Require a non-null ExternalCommandData to avoid nullable conversion warnings and ensure commands
    // always have the data they expect.
    public JSE_RevitAddin_MEP_OPENINGSViewModel(ExternalCommandData commandData)
        {
            _commandData = commandData;
            PlaceOpeningsCommand = new RelayCommand(ExecutePlaceOpenings);
            AddMarkParameterCommand = new RelayCommand(ExecuteAddMarkParameter);
        }

        public ICommand PlaceOpeningsCommand { get; }
        public ICommand AddMarkParameterCommand { get; }

        private void ExecutePlaceOpenings()
        {
            var command = new OpeningsPLaceCommand();
            string message = string.Empty;
            ElementSet elements = new ElementSet();
            command.Execute(_commandData, ref message, elements);
        }

        private void ExecuteAddMarkParameter()
        {
            var command = new MarkParameterAddValue();
            string message = string.Empty;
            ElementSet elements = new ElementSet();
            command.Execute(_commandData, ref message, elements);
        }
    }
}
