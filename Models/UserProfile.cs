using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Xml.Serialization;

namespace JSE_RevitAddin_MEP_OPENINGS.Models
{
    /// <summary>
    /// Represents a user profile with discipline information and preferences
    /// </summary>
    [XmlRoot("UserProfile")]
    public class UserProfile : INotifyPropertyChanged
    {
        private Guid _id;
        private string _name = string.Empty;
        private List<Discipline> _disciplines = new List<Discipline>();
        private string _language = "English";
        private DateTime _createdDate;
        private bool _isActive = true;

        public Guid Id
        {
            get => _id;
            set => SetProperty(ref _id, value);
        }

        public string Name
        {
            get => _name;
            set => SetProperty(ref _name, value);
        }

        public List<Discipline> Disciplines
        {
            get => _disciplines;
            set => SetProperty(ref _disciplines, value);
        }

        public string Language
        {
            get => _language;
            set => SetProperty(ref _language, value);
        }

        public DateTime CreatedDate
        {
            get => _createdDate;
            set => SetProperty(ref _createdDate, value);
        }

        public bool IsActive
        {
            get => _isActive;
            set => SetProperty(ref _isActive, value);
        }
        
        /// <summary>
        /// User's configuration settings for the application
        /// </summary>
        public UserConfiguration? Configuration { get; set; }
        
        /// <summary>
        /// When this profile was last modified
        /// </summary>
        public DateTime LastModified { get; set; } = DateTime.Now;

        /// <summary>
        /// Gets the primary discipline (only one allowed)
        /// </summary>
        public Discipline? PrimaryDiscipline => Disciplines?.FirstOrDefault(d => d.IsPrimary);

        /// <summary>
        /// Gets all MEP disciplines (can be multiple)
        /// </summary>
        public List<Discipline> MepDisciplines => Disciplines?.Where(d => !d.IsPrimary && d.IsSelected).ToList() ?? new List<Discipline>();

        /// <summary>
        /// Gets a display name for the profile
        /// </summary>
        public string DisplayName => $"{Name} ({PrimaryDiscipline?.Name ?? "No Primary Discipline"})";

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
    }
}
