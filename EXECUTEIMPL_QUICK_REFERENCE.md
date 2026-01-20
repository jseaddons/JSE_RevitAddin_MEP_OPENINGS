# ExecuteImpl Pattern - Quick Reference Card

## 🚨 **CRITICAL PATTERN FOR REVIT COMMAND EXECUTION**

### ❌ **NEVER DO THIS**
```csharp
// ❌ ExternalCommandData cannot be constructed
var commandData = new ExternalCommandData(); // CS1729 Error
var commandData = new ExternalCommandData(app, view); // CS1729 Error
```

### ✅ **ALWAYS DO THIS**
```csharp
// ✅ Use ExecuteImpl pattern for direct execution
if (command is DuctSleeveCommand dsc)
{
    var result = dsc.ExecuteImpl(_uiDocument.Application);
}
```

---

## 🏗️ **Command Implementation Template**

```csharp
[Transaction(TransactionMode.Manual)]
public class YourCommand : IExternalCommand
{
    // ✅ Standard Revit interface (for Revit command system)
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;
        return ExecuteCore(uidoc, doc, ref message, elements, null);
    }

    // ✅ Direct execution interface (for orchestrator use)
    public Result ExecuteImpl(UIApplication uiApp)
    {
        try
        {
            UIDocument uidoc = uiApp.ActiveUIDocument;
            Document doc = uidoc.Document;
            string message = "";
            ElementSet elements = new ElementSet();
            
            return ExecuteCore(uidoc, doc, ref message, elements, null);
        }
        catch (Exception ex)
        {
            DebugLogger.Error($"[YourCommand] ExecuteImpl failed: {ex.Message}");
            return Result.Failed;
        }
    }

    // ✅ Core logic separated from Revit interface
    private Result ExecuteCore(UIDocument uidoc, Document doc, ref string message, ElementSet elements, object? additionalParams)
    {
        // Your actual business logic goes here
        return Result.Succeeded;
    }
}
```

---

## 🎯 **Orchestrator Integration Template**

```csharp
private CommandExecutionResult ExecuteCommandWithResourceManagement(IExternalCommand command)
{
    var result = new CommandExecutionResult();
    
    try
    {
        if (command is YourCommand yourCmd)
        {
            // Pass any required parameters
            var commandResult = yourCmd.ExecuteImpl(_uiDocument.Application);
            result.Success = commandResult == Result.Succeeded;
            result.Message = "YourCommand executed";
        }
        else
        {
            DebugLogger.Warning($"Command {command.GetType().Name} does not support direct execution - skipping");
            result.Success = false;
            result.ErrorMessage = "Command does not support direct execution";
        }
    }
    catch (Exception ex)
    {
        result.Success = false;
        result.ErrorMessage = ex.Message;
    }
    
    return result;
}
```

---

## 🔑 **Key Points**

1. **ExternalCommandData is Revit-Only** - Never construct it manually
2. **ExecuteImpl takes UIApplication** - Provides all necessary Revit context
3. **ExecuteCore contains logic** - Separated from Revit interface
4. **Error handling is critical** - Always wrap in try-catch
5. **Backward compatibility** - Commands still implement IExternalCommand

---

## 📚 **Full Documentation**

- **📖 Complete Guide**: [External Command Execution Pattern](EXTERNAL_COMMAND_EXECUTION_PATTERN.md)
- **🔧 Implementation**: [Clearance Backend Implementation Guide](CLEARANCE_BACKEND_IMPLEMENTATION_GUIDE.md)

---

**⚠️ CRITICAL: This pattern is essential for the application to function correctly.**
