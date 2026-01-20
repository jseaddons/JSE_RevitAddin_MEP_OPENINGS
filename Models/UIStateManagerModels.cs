using System;
using System.ComponentModel;

namespace JSE_RevitAddin_MEP_OPENINGS.Models
{
    /// <summary>
    /// Represents the current state of opening conditions in the UI
    /// </summary>
    public class OpeningConditionsState : INotifyPropertyChanged
    {
        private double _ductClearance = 50.0;
        private double _pipeClearance = 75.0;
        private double _cableTrayClearance = 25.0;
        private double _oversizeValue = 0.0;
        private double _levelAdjustment = 0.0;
        private bool _isModified = false;
        private DateTime _lastModified = DateTime.Now;

        public double DuctClearance
        {
            get => _ductClearance;
            set
            {
                if (_ductClearance != value)
                {
                    _ductClearance = value;
                    _isModified = true;
                    _lastModified = DateTime.Now;
                    OnPropertyChanged(nameof(DuctClearance));
                    OnPropertyChanged(nameof(IsModified));
                }
            }
        }

        public double PipeClearance
        {
            get => _pipeClearance;
            set
            {
                if (_pipeClearance != value)
                {
                    _pipeClearance = value;
                    _isModified = true;
                    _lastModified = DateTime.Now;
                    OnPropertyChanged(nameof(PipeClearance));
                    OnPropertyChanged(nameof(IsModified));
                }
            }
        }

        public double CableTrayClearance
        {
            get => _cableTrayClearance;
            set
            {
                if (_cableTrayClearance != value)
                {
                    _cableTrayClearance = value;
                    _isModified = true;
                    _lastModified = DateTime.Now;
                    OnPropertyChanged(nameof(CableTrayClearance));
                    OnPropertyChanged(nameof(IsModified));
                }
            }
        }

        public double OversizeValue
        {
            get => _oversizeValue;
            set
            {
                if (_oversizeValue != value)
                {
                    _oversizeValue = value;
                    _isModified = true;
                    _lastModified = DateTime.Now;
                    OnPropertyChanged(nameof(OversizeValue));
                    OnPropertyChanged(nameof(IsModified));
                }
            }
        }

        public double LevelAdjustment
        {
            get => _levelAdjustment;
            set
            {
                if (_levelAdjustment != value)
                {
                    _levelAdjustment = value;
                    _isModified = true;
                    _lastModified = DateTime.Now;
                    OnPropertyChanged(nameof(LevelAdjustment));
                    OnPropertyChanged(nameof(IsModified));
                }
            }
        }

        public bool IsModified
        {
            get => _isModified;
            private set
            {
                if (_isModified != value)
                {
                    _isModified = value;
                    OnPropertyChanged(nameof(IsModified));
                }
            }
        }

        public DateTime LastModified
        {
            get => _lastModified;
            private set
            {
                if (_lastModified != value)
                {
                    _lastModified = value;
                    OnPropertyChanged(nameof(LastModified));
                }
            }
        }

        public void MarkAsSaved()
        {
            _isModified = false;
            OnPropertyChanged(nameof(IsModified));
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    /// <summary>
    /// Represents the current state of configuration settings in the UI
    /// </summary>
    public class ConfigurationSettingsState : INotifyPropertyChanged
    {
        private double _minOpeningSize = 0.1;
        private double _maxOpeningSize = 1000.0;
        private bool _includeInsulation = true;
        private bool _autoUpdateOpenings = false;
        private bool _isModified = false;
        private DateTime _lastModified = DateTime.Now;

        public double MinOpeningSize
        {
            get => _minOpeningSize;
            set
            {
                if (_minOpeningSize != value)
                {
                    _minOpeningSize = value;
                    _isModified = true;
                    _lastModified = DateTime.Now;
                    OnPropertyChanged(nameof(MinOpeningSize));
                    OnPropertyChanged(nameof(IsModified));
                }
            }
        }

        public double MaxOpeningSize
        {
            get => _maxOpeningSize;
            set
            {
                if (_maxOpeningSize != value)
                {
                    _maxOpeningSize = value;
                    _isModified = true;
                    _lastModified = DateTime.Now;
                    OnPropertyChanged(nameof(MaxOpeningSize));
                    OnPropertyChanged(nameof(IsModified));
                }
            }
        }

        public bool IncludeInsulation
        {
            get => _includeInsulation;
            set
            {
                if (_includeInsulation != value)
                {
                    _includeInsulation = value;
                    _isModified = true;
                    _lastModified = DateTime.Now;
                    OnPropertyChanged(nameof(IncludeInsulation));
                    OnPropertyChanged(nameof(IsModified));
                }
            }
        }

        public bool AutoUpdateOpenings
        {
            get => _autoUpdateOpenings;
            set
            {
                if (_autoUpdateOpenings != value)
                {
                    _autoUpdateOpenings = value;
                    _isModified = true;
                    _lastModified = DateTime.Now;
                    OnPropertyChanged(nameof(AutoUpdateOpenings));
                    OnPropertyChanged(nameof(IsModified));
                }
            }
        }

        public bool IsModified
        {
            get => _isModified;
            private set
            {
                if (_isModified != value)
                {
                    _isModified = value;
                    OnPropertyChanged(nameof(IsModified));
                }
            }
        }

        public DateTime LastModified
        {
            get => _lastModified;
            private set
            {
                if (_lastModified != value)
                {
                    _lastModified = value;
                    OnPropertyChanged(nameof(LastModified));
                }
            }
        }

        public void MarkAsSaved()
        {
            _isModified = false;
            OnPropertyChanged(nameof(IsModified));
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        /// <summary>
        /// Test method to verify the model works correctly
        /// </summary>
        public static void TestModel()
        {
            var model = new OpeningConditionsState();

            // Test initial values
            Console.WriteLine($"Initial values: Duct={model.DuctClearance}, Pipe={model.PipeClearance}, IsModified={model.IsModified}");

            // Test property changes
            model.DuctClearance = 100.0;
            Console.WriteLine($"After change: Duct={model.DuctClearance}, IsModified={model.IsModified}");

            // Test MarkAsSaved
            model.MarkAsSaved();
            Console.WriteLine($"After MarkAsSaved: IsModified={model.IsModified}");

            Console.WriteLine("OpeningConditionsState test completed successfully!");
        }
    }

    /// <summary>
    /// Static test class for ConfigurationSettingsState
    /// </summary>
    public static class ConfigurationSettingsStateTest
    {
        /// <summary>
        /// Test method to verify the model works correctly
        /// </summary>
        public static void TestModel()
        {
            var model = new ConfigurationSettingsState();

            // Test initial values
            Console.WriteLine($"Initial values: MinSize={model.MinOpeningSize}, IncludeInsulation={model.IncludeInsulation}, IsModified={model.IsModified}");

            // Test property changes
            model.IncludeInsulation = false;
            Console.WriteLine($"After change: IncludeInsulation={model.IncludeInsulation}, IsModified={model.IsModified}");

            // Test MarkAsSaved
            model.MarkAsSaved();
            Console.WriteLine($"After MarkAsSaved: IsModified={model.IsModified}");

            Console.WriteLine("ConfigurationSettingsState test completed successfully!");
        }
    }
}
