# ConVoid-like Features Implementation Plan

This document outlines the plan to implement the backend logic for ConVoid-like advanced settings, as described in the provided URL.

## 1. Prioritized Features (Urgent)

The following features have been identified as high priority and will be implemented first.

*   **Limits Section:**
    *   `ignore_openings_smaller_than`
    *   `round_to_rectangular_if_diameter_greater_than`
    *   `join_openings_at_distance`
    *   `ignore_openings_with_angle_greater_than`
    *   `round_up_opening_dimensions`

## 2. Greyed Out Features (Future Extensions)

The following features are planned for future releases and their corresponding UI elements will be greyed out in the current implementation.

*   **Manage Section:**
    *   `reset_approval_status`
    *   `dimension_change_threshold`
    *   `location_change_threshold`
*   **Elements Section:**
    *   `cut_openings_with_hosts`
    *   `create_constraint_between_openings_and_hosts`
    *   `adopt_provision_for_voids`
*   **Element Filter Section:**
    *   `include_host_elements_not_in_view`
    *   `include_reference_elements_not_in_view`
    *   `include_host_elements_in_demolished_phase`
*   **Limits Section:**
    *   `create_openings_with_slope`

## 3. Data Model Integration

### 3.1. Update `UserConfiguration`

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

## 4. Backend Logic Implementation

### 4.1. Create `SettingsService`

A new service will be created to manage the loading and saving of the `SettingsModel` for a user profile.

- **Action:** Create a new file `Services/SettingsService.cs`.

### 4.2. Integrate `SettingsService` into `ApplicationProfileService`

The `ApplicationProfileService` will use the `SettingsService` to manage the settings.

- **Action:** Update `Services/ApplicationProfileService.cs`.

### 4.3. Implement Logic for Each Setting

The logic for each setting will be implemented in the relevant services, with priority given to the features listed in the "Prioritized Features" section.

## 5. UI Modifications

The user has indicated that the UI is already implemented. We will connect the UI to the `SettingsModel` properties in the corresponding ViewModels to ensure that the settings are correctly displayed and saved. We will also disable the UI controls for the features listed in the "Greyed Out Features" section.
