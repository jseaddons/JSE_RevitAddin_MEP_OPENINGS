using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Views;

namespace JSE_RevitAddin_MEP_OPENINGS.ViewModels
{
    /// <summary>
    /// ViewModel for the Profile Management Dialog
    /// </summary>
    public partial class ProfileManagementViewModel : ObservableObject
    {
        private readonly ApplicationProfileService _appProfileService;
        private UserProfile? _selectedProfile;

        [ObservableProperty]
        private string _currentUserInfo = string.Empty;

        [ObservableProperty]
        private List<UserProfile> _availableProfiles = new List<UserProfile>();

        public UserProfile? SelectedProfile
        {
            get => _selectedProfile;
            set 
            { 
                if (SetProperty(ref _selectedProfile, value))
                {
                    SwitchProfileCommand.NotifyCanExecuteChanged();
                }
            }
        }

        public event EventHandler? DialogClosed;
        public event EventHandler<ProfileCreatedEventArgs>? ProfileCreated;

        public ProfileManagementViewModel(ApplicationProfileService appProfileService)
        {
            _appProfileService = appProfileService ?? throw new ArgumentNullException(nameof(appProfileService));

            LoadData();
        }

        /// <summary>
        /// Command to create a new profile
        /// </summary>
        [RelayCommand]
        private void CreateNewProfile()
        {
            try
            {
                // EMERGENCY MODE: Use WinForms instead of WPF to avoid crashes
                var emergencyDialog = new EmergencyProfileSetup(_appProfileService.ProfileService, _appProfileService.StatusManager);
                
                var result = emergencyDialog.ShowDialog();
                if (result == System.Windows.Forms.DialogResult.OK && emergencyDialog.CreatedProfile != null)
                {
                    // Set the created profile as current
                    _appProfileService.SetCurrentProfile(emergencyDialog.CreatedProfile);

                    // Refresh the data
                    LoadData();

                    ProfileCreated?.Invoke(this, new ProfileCreatedEventArgs(emergencyDialog.CreatedProfile));

                    // Auto-open MainDialog after profile creation
                    OpenMainDialog();
                }
            }
            catch (Exception ex)
            {
                _appProfileService.StatusManager.UpdateStatus($"Failed to create profile: {ex.Message}", StatusType.Error);
            }
        }

        /// <summary>
        /// Command to switch to the selected profile
        /// </summary>
        [RelayCommand(CanExecute = nameof(CanSwitchProfile))]
        private void SwitchProfile()
        {
            if (SelectedProfile == null)
                return;

            try
            {
                _appProfileService.SetCurrentProfile(SelectedProfile);
                LoadData();

                _appProfileService.StatusManager.UpdateStatus(
                    $"Switched to profile '{SelectedProfile.Name}'",
                    StatusType.Success);

                // Auto-open MainDialog after profile switch
                OpenMainDialog();
            }
            catch (Exception ex)
            {
                _appProfileService.StatusManager.UpdateStatus($"Failed to switch profile: {ex.Message}", StatusType.Error);
            }
        }

        /// <summary>
        /// Command to close the dialog
        /// </summary>
        [RelayCommand]
        private void Close()
        {
            DialogClosed?.Invoke(this, EventArgs.Empty);
        }

        /// <summary>
        /// Checks if profile can be switched
        /// </summary>
        private bool CanSwitchProfile()
        {
            return SelectedProfile != null && SelectedProfile.Id != _appProfileService.CurrentProfile?.Id;
        }

        /// <summary>
        /// Loads the current data
        /// </summary>
        private void LoadData()
        {
            CurrentUserInfo = _appProfileService.GetUserInfoString();
            AvailableProfiles = _appProfileService.ProfileService.AvailableProfiles.ToList();

            // Select current profile if available
            if (_appProfileService.CurrentProfile != null)
            {
                SelectedProfile = AvailableProfiles.FirstOrDefault(p => p.Id == _appProfileService.CurrentProfile.Id);
            }
        }

        /// <summary>
        /// Opens the MainDialog with the current profile
        /// </summary>
        private void OpenMainDialog()
        {
            try
            {
                // EMERGENCY MODE: Use WinForms MainDialog instead of WPF to prevent crashes
                var emergencyMainDialog = new EmergencyMainDialog(_appProfileService);
                emergencyMainDialog.Show(); // Modeless for Revit compatibility
            }
            catch (Exception ex)
            {
                _appProfileService.StatusManager.UpdateStatus($"Failed to open MainDialog: {ex.Message}", StatusType.Error);
            }
        }


    }
}
