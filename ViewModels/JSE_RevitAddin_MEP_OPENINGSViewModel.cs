
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
    private readonly ExternalCommandData? _commandData;
    private readonly System.Action<ExternalCommandData>? _placeAction;
    private readonly System.Action<ExternalCommandData>? _addMarkAction;

    public JSE_RevitAddin_MEP_OPENINGSViewModel(ExternalCommandData? commandData = null, System.Action<ExternalCommandData>? placeAction = null, System.Action<ExternalCommandData>? addMarkAction = null)
        {
            _commandData = commandData;
            _placeAction = placeAction;
            _addMarkAction = addMarkAction;
            PlaceOpeningsCommand = new RelayCommand(ExecutePlaceOpenings);
            AddMarkParameterCommand = new RelayCommand(ExecuteAddMarkParameter);
        }

        public ICommand PlaceOpeningsCommand { get; }
        public ICommand AddMarkParameterCommand { get; }

        private void ExecutePlaceOpenings()
        {
            if (_commandData == null) return;
            if (_placeAction != null)
            {
                _placeAction.Invoke(_commandData);
                return;
            }
            // Fallback -- OpeningsPLaceCommand is obsolete
            // All placement should go through SleevePlacementExternalEvent via the Place Openings button
            System.Diagnostics.Debug.WriteLine("Attempted to use obsolete OpeningsPLaceCommand - please use the Place Openings workflow instead");
        }

        private void ExecuteAddMarkParameter()
        {
            if (_commandData == null) return;
            if (_addMarkAction != null)
            {
                _addMarkAction.Invoke(_commandData);
                return;
            }
            // Fallback -- directly execute command if no delegate supplied (legacy behavior)
            var command = new MarkParameterAddValue();
            string? message = null;
            ElementSet elements = new ElementSet();
            command.Execute(_commandData, ref message!, elements);
        }
    }
}
