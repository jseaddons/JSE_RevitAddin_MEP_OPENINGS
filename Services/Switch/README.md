# 🔌 Switch Classes - Quick Reference Guide

## How to Turn Diagnostic Mode ON and OFF

### Method 1: Using MasterSwitch (Recommended)

```csharp
using JSE_RevitAddin_MEP_OPENINGS.Services.Switch;

// ✅ TURN ON (Enable diagnostic mode - full logging)
MasterSwitch.EnableDiagnostics();

// ❌ TURN OFF (Disable diagnostic mode - minimal logging)
MasterSwitch.DisableDiagnostics();

// 🔄 TOGGLE (Switch between on/off)
MasterSwitch.ToggleDiagnostics();
```

### Method 2: Using DiagnosticSwitch Directly

```csharp
using JSE_RevitAddin_MEP_OPENINGS.Services.Switch;

// ✅ TURN ON
DiagnosticSwitch.Enable();

// ❌ TURN OFF
DiagnosticSwitch.Disable();

// 🔄 TOGGLE
DiagnosticSwitch.Toggle();
```

### Method 3: Quick Mode Presets

```csharp
using JSE_RevitAddin_MEP_OPENINGS.Services.Switch;

// 🚀 DEPLOYMENT MODE (OFF - minimal logging, optimizations ON)
MasterSwitch.SetDeploymentMode();

// 🔧 DEVELOPMENT MODE (ON - full logging, optimizations ON)
MasterSwitch.SetDevelopmentMode();

// 🐛 DEBUG MODE (ON - full logging, optimizations OFF)
MasterSwitch.SetDebugMode();
```

## Check Current Status

```csharp
using JSE_RevitAddin_MEP_OPENINGS.Services.Switch;

// Check if diagnostic mode is enabled
bool isEnabled = DiagnosticSwitch.IsEnabled;

// Get formatted status string
string status = DiagnosticSwitch.GetStatus();
// Returns: "✅ Diagnostic Mode: ENABLED" or "⚠️ Diagnostic Mode: DISABLED"

// Get complete status (all switches)
string fullStatus = MasterSwitch.GetCompleteStatus();
```

## Where to Use

You can call these methods from **anywhere** in your code:

1. **In a Command class** (e.g., at the start of `Execute()` method)
2. **In a UI button click handler**
3. **In application startup code**
4. **In a test/debug command**

## Example: Add to a Command

```csharp
using JSE_RevitAddin_MEP_OPENINGS.Services.Switch;

public class MyCommand : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        // Turn diagnostic mode ON
        MasterSwitch.EnableDiagnostics();
        
        // Your code here...
        
        return Result.Succeeded;
    }
}
```

## Example: Toggle in UI

```csharp
private void ToggleDiagnosticButton_Click(object sender, EventArgs e)
{
    MasterSwitch.ToggleDiagnostics();
    
    // Show status to user
    string status = DiagnosticSwitch.GetStatus();
    MessageBox.Show(status);
}
```

## What Each Mode Does

| Mode | Logging | Optimizations | Use Case |
|------|---------|---------------|----------|
| **Diagnostic ON** | ✅ Full | ✅ Enabled | Development, debugging |
| **Diagnostic OFF** | ❌ Minimal | ✅ Enabled | Production, deployment |
| **Debug Mode** | ✅ Full | ❌ Disabled | Troubleshooting issues |

## Quick Reference

```csharp
// Turn ON
MasterSwitch.EnableDiagnostics();

// Turn OFF  
MasterSwitch.DisableDiagnostics();

// Toggle
MasterSwitch.ToggleDiagnostics();

// Check status
bool isOn = DiagnosticSwitch.IsEnabled;
```

