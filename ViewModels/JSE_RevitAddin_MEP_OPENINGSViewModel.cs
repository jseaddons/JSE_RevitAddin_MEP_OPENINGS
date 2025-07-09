
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

        public JSE_RevitAddin_MEP_OPENINGSViewModel(ExternalCommandData commandData = null)
        {
            _commandData = commandData;
            PlaceOpeningsCommand = new RelayCommand(ExecutePlaceOpenings);
            AddMarkParameterCommand = new RelayCommand(ExecuteAddMarkParameter);
        }

        public ICommand PlaceOpeningsCommand { get; }
        public ICommand AddMarkParameterCommand { get; }

        private void ExecutePlaceOpenings()
        {
            if (_commandData == null) return;

            var command = new OpeningsPLaceCommand();
            string message = null;
            ElementSet elements = new ElementSet();
            command.Execute(_commandData, ref message, elements);
        }

        private void ExecuteAddMarkParameter()
        {
            if (_commandData == null) return;

            var command = new MarkParameterAddValue();
            string message = null;
            ElementSet elements = new ElementSet();
            command.Execute(_commandData, ref message, elements);
        }
    }
}