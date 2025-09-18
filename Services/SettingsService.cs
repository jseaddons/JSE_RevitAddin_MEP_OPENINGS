using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    public class SettingsService
    {
        public SettingsModel GetSettings(UserProfile profile)
        {
            // For now, we return the default settings.
            // Later, we will implement loading from a file.
            return profile.Configuration?.AdvancedSettings ?? new SettingsModel();
        }

        public void SaveSettings(UserProfile profile, SettingsModel settings)
        {
            if (profile.Configuration != null)
            {
                profile.Configuration.AdvancedSettings = settings;
                // Later, we will implement saving to a file.
            }
        }
    }
}
