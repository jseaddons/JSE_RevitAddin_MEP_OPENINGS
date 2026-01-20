# ConVoid-like Features Implementation Plan

This document outlines the plan to implement the backend logic for ConVoid-like advanced settings, as described in the provided URL.

## 1. Data Model Integration

### 1.1. Update `UserConfiguration`

The `SettingsModel` class already exists and contains the required properties. We need to integrate it into the user's configuration.

- **Action:** Add a property of type `SettingsModel` to the `UserConfiguration` class in `Models/UserConfiguration.cs`.

```csharp
// In Models/UserConfiguration.cs

public class UserConfiguration
{
    // ... existing properties

    /// <summary>
    /// Advanced settings for ConVoid-like features.
    /// </summary>
    public SettingsModel AdvancedSettings { get; set; } = new SettingsModel();
}
```

## 2. Backend Logic Implementation

### 2.1. Create `SettingsService`

A new service will be created to manage the loading and saving of the `SettingsModel` for a user profile.

- **Action:** Create a new file `Services/SettingsService.cs`.
- **Content:**

```csharp
// In Services/SettingsService.cs

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
```

### 2.2. Integrate `SettingsService` into `ApplicationProfileService`

The `ApplicationProfileService` will use the `SettingsService` to manage the settings.

- **Action:** Update `Services/ApplicationProfileService.cs`.
- **Content:**

```csharp
// In Services/ApplicationProfileService.cs

public class ApplicationProfileService
{
    // ... existing properties and methods

    private readonly SettingsService _settingsService;

    public ApplicationProfileService(ProfileManagementService profileService, StatusManager statusManager)
    {
        // ... existing code
        _settingsService = new SettingsService();
    }

    public SettingsModel GetCurrentSettings()
    {
        if (CurrentProfile == null)
            return new SettingsModel();

        return _settingsService.GetSettings(CurrentProfile);
    }

    public void SaveCurrentSettings(SettingsModel settings)
    {
        if (CurrentProfile != null)
        {
            _settingsService.SaveSettings(CurrentProfile, settings);
        }
    }
}
```

### 2.3. Implement Logic for Each Setting

The logic for each setting will be implemented in the relevant services.

-   **`reset_approval_status`**: In `OpeningTrackingService`, if an opening is found to be modified, and this setting is true, set the approval status to "pending".
-   **`dimension_change_threshold`**: In `OpeningTrackingService`, use this value to check if the dimensions of an opening have changed significantly.
-   **`location_change_threshold`**: In `OpeningTrackingService`, use this value to check if the location of an opening have changed significantly.
-   **`cut_openings_with_hosts`**: In `PipeSleevePlacer`, `DuctSleevePlacer`, and `CableTraySleevePlacer`, use this setting to decide whether to cut the host element when placing an opening.
-   **`create_constraint_between_openings_and_hosts`**: In `PipeSleevePlacer`, `DuctSleevePlacer`, and `CableTraySleevePlacer`, use this setting to decide whether to create a constraint between the opening and the host element.
-   **`adopt_provision_for_voids`**: In `LinkedFileService`, use this setting to link adopted openings with the original openings in the linked file.
-   **`include_host_elements_not_in_view`**: In `MepElementCollectorHelper`, use this setting to filter host elements based on their visibility.
-   **`include_reference_elements_not_in_view`**: In `MepElementCollectorHelper`, use this setting to filter reference elements based on their visibility.
-   **`include_host_elements_in_demolished_phase`**: In `MepElementCollectorHelper`, use this setting to include or exclude demolished host elements.
-   **`ignore_openings_smaller_than`**: In `DuctSleevePlacerService`, `PipeSleevePlacerService`, and `CableTraySleevePlacerService`, use this setting to ignore small openings.
-   **`round_to_rectangular_if_diameter_greater_than`**: In `DuctSleevePlacerService`, `PipeSleevePlacerService`, and `CableTraySleevePlacerService`, use this setting to convert round openings to rectangular.
-   **`join_openings_at_distance`**: In `ClusterSleeveDuplicationService`, use this setting to join nearby openings.
-   **`ignore_openings_with_angle_greater_than`**: In `WallIntersectionService` and `EfficientIntersectionService`, use this setting to ignore openings with a steep intersection angle.
-   **`round_up_opening_dimensions`**: In `DuctSleevePlacerService`, `PipeSleevePlacerService`, and `CableTraySleevePlacerService`, use this setting to round up opening dimensions.

## 3. UI Modifications

The user has indicated that the UI is already implemented. We will connect the UI to the `SettingsModel` properties in the corresponding ViewModels to ensure that the settings are correctly displayed and saved. We will also disable the UI controls for the features that are not to be implemented.

## 4. Future Extensions

The following features will be ignored for now, and the corresponding UI elements will be greyed out:

-   Manage
-   Elements
-   Elements Filter
-   Create openings with slope (in Limits)