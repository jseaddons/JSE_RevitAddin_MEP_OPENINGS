using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace JSE_RevitAddin_MEP_OPENINGS.Models
{
    /// <summary>
    /// Represents a discipline with validation rules and allowed combinations
    /// </summary>
    public class Discipline : INotifyPropertyChanged
    {
        private string _name = string.Empty;
        private bool _isPrimary = false;
        private bool _isSelected = false;
        private List<string> _allowedCombinations = new List<string>();
        private string _description = string.Empty;

        public Discipline()
        {
        }

        public Discipline(string name, bool isPrimary, string description = "")
        {
            _name = name;
            _isPrimary = isPrimary;
            _description = description;
        }

        public string Name
        {
            get => _name;
            set => SetProperty(ref _name, value);
        }

        public bool IsPrimary
        {
            get => _isPrimary;
            set => SetProperty(ref _isPrimary, value);
        }

        public bool IsSelected
        {
            get => _isSelected;
            set => SetProperty(ref _isSelected, value);
        }

        public List<string> AllowedCombinations
        {
            get => _allowedCombinations;
            set => SetProperty(ref _allowedCombinations, value);
        }

        public string Description
        {
            get => _description;
            set => SetProperty(ref _description, value);
        }

        /// <summary>
        /// Gets the discipline type (Primary or MEP)
        /// </summary>
        public DisciplineType Type => IsPrimary ? DisciplineType.Primary : DisciplineType.Mep;

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        protected bool SetProperty<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value)) return false;
            field = value;
            OnPropertyChanged(propertyName);
            return true;
        }

        public override string ToString()
        {
            return Name;
        }

        public override bool Equals(object? obj)
        {
            if (obj is Discipline other)
            {
                return Name.Equals(other.Name, StringComparison.OrdinalIgnoreCase);
            }
            return false;
        }

        public override int GetHashCode()
        {
            return Name.ToUpperInvariant().GetHashCode();
        }
    }

    /// <summary>
    /// Discipline types for validation
    /// </summary>
    public enum DisciplineType
    {
        Primary,
        Mep
    }
}
