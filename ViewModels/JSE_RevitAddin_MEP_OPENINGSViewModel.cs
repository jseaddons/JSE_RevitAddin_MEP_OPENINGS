
using System.Windows.Input;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.ComponentModel;
using System.Windows.Input;
using JSE_RevitAddin_MEP_OPENINGS.Commands;

namespace JSE_RevitAddin_MEP_OPENINGS.ViewModels
{
    public sealed class JSE_RevitAddin_MEP_OPENINGSViewModel : INotifyPropertyChanged
    {
    private readonly ExternalCommandData? _commandData;
    private readonly System.Action<ExternalCommandData>? _placeAction;
    private readonly System.Action<ExternalCommandData>? _addMarkAction;

    public JSE_RevitAddin_MEP_OPENINGSViewModel(ExternalCommandData? commandData = null, System.Action<ExternalCommandData>? placeAction = null, System.Action<ExternalCommandData>? addMarkAction = null)
        {
            _commandData = commandData;
            _placeAction = placeAction;
            _addMarkAction = addMarkAction;
            PlaceOpeningsCommand = new SimpleCommand(ExecutePlaceOpenings);
            AddMarkParameterCommand = new SimpleCommand(ExecuteAddMarkParameter);
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
            // Note: MarkParameterAddValue moved to Backup folder - using new MarkParameterCommand instead
            System.Diagnostics.Debug.WriteLine("[ViewModel] MarkParameterAddValue fallback not available - command moved to Backup folder");
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    // Simple command implementation to replace RelayCommand
    public class SimpleCommand : ICommand
    {
        private readonly Action _execute;
        private readonly Func<bool>? _canExecute;

        public SimpleCommand(Action execute, Func<bool>? canExecute = null)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }

        public event EventHandler? CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public bool CanExecute(object? parameter)
        {
            return _canExecute?.Invoke() ?? true;
        }

        public void Execute(object? parameter)
        {
            _execute();
        }
    }
}
