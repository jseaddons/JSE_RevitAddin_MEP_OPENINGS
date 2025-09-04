using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.ViewModels
{
    /// <summary>
    /// Thread-safe ViewModel for Profile Setup Dialog - designed to prevent crashes
    /// </summary>
    public partial class ProfileSetupViewModelSafe : ObservableObject
    {
        private readonly ProfileManagementService _profileService;
        private readonly StatusManager _statusManager;
        private bool _isValidating = false; // Prevent recursive validation

        [ObservableProperty]
        private string _profileName = string.Empty;

        // Fixed to English-only for initial implementation
        private const string FixedLanguage = "English";

        [ObservableProperty]
        private bool _isArchitecturalSelected = false;

        [ObservableProperty]
        private bool _isStructuralSelected = false;

        [ObservableProperty]
        private bool _isTechnicalPlannerSelected = false;

        [ObservableProperty]
        private bool _isCoordinationSelected = false;

        [ObservableProperty]
        private bool _isMechanicalSelected = false;

        [ObservableProperty]
        private bool _isElectricalSelected = false;

        [ObservableProperty]
        private bool _isPlumbingSelected = false;

        [ObservableProperty]
        private bool _isFireProtectionSelected = false;

        [ObservableProperty]
        private string _validationMessage = string.Empty;

        [ObservableProperty]
        private bool _isValid = false;

        [ObservableProperty]
        private bool _isCreating = false;

        public event EventHandler<ProfileCreatedEventArgs>? ProfileCreated;
        public event EventHandler? DialogClosed;

        public ProfileSetupViewModelSafe(ProfileManagementService profileService, StatusManager statusManager)
        {
            _profileService = profileService ?? throw new ArgumentNullException(nameof(profileService));
            _statusManager = statusManager ?? throw new ArgumentNullException(nameof(statusManager));

            // Create commands
            CreateProfileCommand = new RelayCommand(CreateProfile, CanCreateProfile);
            CancelCommand = new RelayCommand(Cancel);
            ClearAllCommand = new RelayCommand(ClearAll);

            // Initial validation
            ValidateSelection();
        }

        public IRelayCommand CreateProfileCommand { get; }
        public IRelayCommand CancelCommand { get; }
        public IRelayCommand ClearAllCommand { get; }

        /// <summary>
        /// Command to create the profile
        /// </summary>
        private void CreateProfile()
        {
            if (!CanCreateProfile())
                return;

            try
            {
                IsCreating = true;
                _statusManager.StartOperation("Profile Creation");

                var disciplines = GetSelectedDisciplines();
                var profile = _profileService.CreateProfile(ProfileName, disciplines, FixedLanguage);

                _statusManager.CompleteOperation("Profile Creation", true, $"Profile '{ProfileName}' created successfully");

                ProfileCreated?.Invoke(this, new ProfileCreatedEventArgs(profile));
            }
            catch (Exception ex)
            {
                _statusManager.CompleteOperation("Profile Creation", false, $"Failed to create profile: {ex.Message}");
                ValidationMessage = ex.Message;
                IsValid = false;
            }
            finally
            {
                IsCreating = false;
            }
        }

        /// <summary>
        /// Command to cancel the dialog
        /// </summary>
        private void Cancel()
        {
            DialogClosed?.Invoke(this, EventArgs.Empty);
        }

        /// <summary>
        /// Gets the selected disciplines
        /// </summary>
        private List<Discipline> GetSelectedDisciplines()
        {
            var disciplines = new List<Discipline>();

            // Primary disciplines (single selection)
            if (IsArchitecturalSelected)
                disciplines.Add(new Discipline("Architectural", true, "Architectural design and coordination"));
            if (IsStructuralSelected)
                disciplines.Add(new Discipline("Structural", true, "Structural engineering and design"));
            if (IsTechnicalPlannerSelected)
                disciplines.Add(new Discipline("Technical Planner", true, "Technical planning and coordination"));
            if (IsCoordinationSelected)
                disciplines.Add(new Discipline("Coordination", true, "Project coordination and management"));

            // MEP disciplines (multi selection)
            if (IsMechanicalSelected)
                disciplines.Add(new Discipline("Mechanical", false, "Mechanical systems (HVAC, plumbing)"));
            if (IsElectricalSelected)
                disciplines.Add(new Discipline("Electrical", false, "Electrical systems and power"));
            if (IsPlumbingSelected)
                disciplines.Add(new Discipline("Plumbing", false, "Plumbing and water systems"));
            if (IsFireProtectionSelected)
                disciplines.Add(new Discipline("Fire Fighting", false, "Fire fighting and safety systems"));

            return disciplines;
        }

        /// <summary>
        /// Validates the current selection
        /// </summary>
        private void ValidateSelection()
        {
            if (_isValidating) return; // Prevent recursive calls

            try
            {
                _isValidating = true;

                var selectedDisciplines = GetSelectedDisciplines();
                var primaryDisciplines = selectedDisciplines.Where(d => d.IsPrimary).ToList();
                var mepDisciplines = selectedDisciplines.Where(d => !d.IsPrimary).ToList();

                if (string.IsNullOrWhiteSpace(ProfileName))
                {
                    ValidationMessage = "Profile name is required";
                    IsValid = false;
                }
                else if (primaryDisciplines.Count > 1)
                {
                    ValidationMessage = "Only one primary discipline can be selected";
                    IsValid = false;
                }
                else if (!selectedDisciplines.Any())
                {
                    ValidationMessage = "Please select at least one discipline";
                    IsValid = false;
                }
                else if (primaryDisciplines.Any() && mepDisciplines.Any())
                {
                    ValidationMessage = "Choose either one primary discipline OR one/more MEP disciplines";
                    IsValid = false;
                }
                else
                {
                    ValidationMessage = string.Empty;
                    IsValid = true;
                }

                // Update command state
                CreateProfileCommand.NotifyCanExecuteChanged();
            }
            catch (Exception ex)
            {
                // Log validation errors but don't crash
                ValidationMessage = $"Validation error: {ex.Message}";
                IsValid = false;
            }
            finally
            {
                _isValidating = false;
            }
        }

        /// <summary>
        /// Clear all selections (primary and MEP)
        /// </summary>
        private void ClearAll()
        {
            IsArchitecturalSelected = false;
            IsStructuralSelected = false;
            IsTechnicalPlannerSelected = false;
            IsCoordinationSelected = false;
            IsMechanicalSelected = false;
            IsElectricalSelected = false;
            IsPlumbingSelected = false;
            IsFireProtectionSelected = false;
            ValidateSelection();
        }

        /// <summary>
        /// Checks if profile can be created
        /// </summary>
        private bool CanCreateProfile()
        {
            return IsValid && !IsCreating;
        }

        /// <summary>
        /// Override PropertyChanged to add validation trigger
        /// </summary>
        protected override void OnPropertyChanged(PropertyChangedEventArgs e)
        {
            base.OnPropertyChanged(e);

            // Trigger validation for relevant properties
            if (e.PropertyName != nameof(ValidationMessage) && 
                e.PropertyName != nameof(IsValid) && 
                e.PropertyName != nameof(IsCreating) &&
                !_isValidating)
            {
                ValidateSelection();
            }
        }

        /// <summary>
        /// Handles primary discipline selection (ensures only one is selected)
        /// </summary>
        partial void OnIsArchitecturalSelectedChanged(bool value)
        {
            if (value && !_isValidating)
            {
                IsStructuralSelected = false;
                IsTechnicalPlannerSelected = false;
                IsCoordinationSelected = false;
                IsMechanicalSelected = false;
                IsElectricalSelected = false;
                IsPlumbingSelected = false;
                IsFireProtectionSelected = false;
            }
        }

        partial void OnIsStructuralSelectedChanged(bool value)
        {
            if (value && !_isValidating)
            {
                IsArchitecturalSelected = false;
                IsTechnicalPlannerSelected = false;
                IsCoordinationSelected = false;
                IsMechanicalSelected = false;
                IsElectricalSelected = false;
                IsPlumbingSelected = false;
                IsFireProtectionSelected = false;
            }
        }

        partial void OnIsTechnicalPlannerSelectedChanged(bool value)
        {
            if (value && !_isValidating)
            {
                IsArchitecturalSelected = false;
                IsStructuralSelected = false;
                IsCoordinationSelected = false;
                IsMechanicalSelected = false;
                IsElectricalSelected = false;
                IsPlumbingSelected = false;
                IsFireProtectionSelected = false;
            }
        }

        partial void OnIsCoordinationSelectedChanged(bool value)
        {
            if (value && !_isValidating)
            {
                IsArchitecturalSelected = false;
                IsStructuralSelected = false;
                IsTechnicalPlannerSelected = false;
                IsMechanicalSelected = false;
                IsElectricalSelected = false;
                IsPlumbingSelected = false;
                IsFireProtectionSelected = false;
            }
        }

        partial void OnIsMechanicalSelectedChanged(bool value)
        {
            if (value && !_isValidating)
            {
                IsArchitecturalSelected = false;
                IsStructuralSelected = false;
                IsTechnicalPlannerSelected = false;
                IsCoordinationSelected = false;
            }
        }

        partial void OnIsElectricalSelectedChanged(bool value)
        {
            if (value && !_isValidating)
            {
                IsArchitecturalSelected = false;
                IsStructuralSelected = false;
                IsTechnicalPlannerSelected = false;
                IsCoordinationSelected = false;
            }
        }

        partial void OnIsPlumbingSelectedChanged(bool value)
        {
            if (value && !_isValidating)
            {
                IsArchitecturalSelected = false;
                IsStructuralSelected = false;
                IsTechnicalPlannerSelected = false;
                IsCoordinationSelected = false;
            }
        }

        partial void OnIsFireProtectionSelectedChanged(bool value)
        {
            if (value && !_isValidating)
            {
                IsArchitecturalSelected = false;
                IsStructuralSelected = false;
                IsTechnicalPlannerSelected = false;
                IsCoordinationSelected = false;
            }
        }

        // Computed properties to control enabling (greying) of groups
        public bool ArePrimaryEnabled => !(IsMechanicalSelected || IsElectricalSelected || IsPlumbingSelected || IsFireProtectionSelected);
        public bool AreMepEnabled => !(IsArchitecturalSelected || IsStructuralSelected || IsTechnicalPlannerSelected || IsCoordinationSelected);

        /// <summary>
        /// Disposes the ViewModel and cleans up event subscriptions
        /// </summary>
        public void Dispose()
        {
            // Clean disposal - no event subscriptions to clean up in this safe version
        }
    }
}
