# Revit ExternalEvent & UI Integration: Step-by-Step Teaching Manual

## Overview
This manual teaches how to safely interact with the Revit API from a custom UI (WinForms/WPF) using `ExternalEvent`, with a practical example: selecting elements in Revit from a modeless UI.

---

## 1. Why Use ExternalEvent?
- **Revit API is not thread-safe**: You cannot call it directly from your UI thread.
- **ExternalEvent** lets you trigger Revit API code from your UI, safely and without freezing or crashing Revit.

---

## 2. Architecture Diagram

```
[User clicks button in UI] 
        ↓
[UI calls ExternalEvent.Raise()]
        ↓
[Revit calls IExternalEventHandler.Execute()]
        ↓
[Your Revit API code runs safely]
```

---

## 3. Step-by-Step Implementation

### Step 1: Implement IExternalEventHandler

```csharp
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;

public class SelectElementHandler : IExternalEventHandler
{
    public void Execute(UIApplication app)
    {
        UIDocument uidoc = app.ActiveUIDocument;
        // Prompt user to select elements
        var refs = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element, "Select elements");
        TaskDialog.Show("Selection", $"You selected {refs.Count} elements.");
    }

    public string GetName() => "Select Element Handler";
}
```

---

#### Detailed Explanation (Line by Line)

```csharp
public class SelectElementHandler : IExternalEventHandler // 1. Define a handler class for ExternalEvent
{
    public void Execute(UIApplication app) // 2. This method is called by Revit when the event is raised
    {
        UIDocument uidoc = app.ActiveUIDocument; // 3. Get the active document (the one the user is working on)
        // Prompt user to select elements
        var refs = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element, "Select elements"); // 4. Open a selection box in Revit for the user to pick elements
        TaskDialog.Show("Selection", $"You selected {refs.Count} elements."); // 5. Show a dialog with the number of selected elements
    }

    public string GetName() => "Select Element Handler"; // 6. Returns the name of this handler (for logging/debugging)
}
```

- **1.** Declares a class that implements `IExternalEventHandler` (required for Revit ExternalEvent).
- **2.** Implements the `Execute` method, which is called on the Revit main thread.
- **3.** Gets the current active document (so you can interact with the model).
- **4.** Prompts the user to select elements in the Revit UI (safe, runs in Revit context).
- **5.** Shows a message box with the count of selected elements.
- **6.** Implements `GetName`, which is used by Revit for identification/logging.

---

### Step 2: Create the ExternalEvent and Pass to UI

```csharp
// In your ExternalCommand or App startup:
var handler = new SelectElementHandler();
var extEvent = ExternalEvent.Create(handler);
var myForm = new MyForm(extEvent);
myForm.Show();
```

---

#### Detailed Explanation (Step 2: Create the ExternalEvent and Pass to UI)

```csharp
// In your ExternalCommand or App startup:
var handler = new SelectElementHandler(); // 1. Create an instance of your handler class
var extEvent = ExternalEvent.Create(handler); // 2. Create the ExternalEvent, passing your handler
var myForm = new MyForm(extEvent); // 3. Create your UI form, passing the ExternalEvent
myForm.Show(); // 4. Show the modeless form (non-blocking)
```

- **1.** Instantiates your handler, which contains the Revit API logic.
- **2.** Creates an ExternalEvent object, which is the bridge between your UI and the Revit API.
- **3.** Passes the ExternalEvent to your UI form so it can raise the event when needed.
- **4.** Shows the form as modeless, so Revit remains responsive.

---

### Step 3: Modeless UI Example (WinForms)

```csharp
using System.Windows.Forms;
using Autodesk.Revit.UI;

public partial class MyForm : Form
{
    private readonly ExternalEvent _extEvent; // 1. Store the ExternalEvent reference

    public MyForm(ExternalEvent extEvent)
    {
        InitializeComponent(); // 2. Standard WinForms setup
        _extEvent = extEvent; // 3. Save the ExternalEvent for later use
    }

    private void selectButton_Click(object sender, EventArgs e)
    {
        _extEvent.Raise(); // 4. When the button is clicked, raise the event to trigger your handler
    }
}
```

---

#### Detailed Explanation (Step 3: Modeless UI Example)

```csharp
using System.Windows.Forms;
using Autodesk.Revit.UI;

public partial class MyForm : Form
{
    private readonly ExternalEvent _extEvent; // 1. Store the ExternalEvent reference

    public MyForm(ExternalEvent extEvent)
    {
        InitializeComponent(); // 2. Standard WinForms setup
        _extEvent = extEvent; // 3. Save the ExternalEvent for later use
    }

    private void selectButton_Click(object sender, EventArgs e)
    {
        _extEvent.Raise(); // 4. When the button is clicked, raise the event to trigger your handler
    }
}
```

- **1.** Keeps a reference to the ExternalEvent so the UI can trigger it.
- **2.** Standard WinForms initialization (sets up controls, etc.).
- **3.** Stores the ExternalEvent passed from the main command.
- **4.** When the user clicks the button, calls `Raise()` to safely run Revit API code.

---

## 4. Full Example: Putting It All Together

- **Handler**: Implements the Revit logic (element selection)
- **ExternalEvent**: Bridges UI and Revit
- **UI**: Calls `Raise()` on button click

---

## 5. Best Practices
- Never call Revit API directly from your UI thread.
- Use `ExternalEvent` for all Revit API calls from modeless UIs.
- Pass data to the handler via fields/properties if needed.
- Use try/catch in your handler for robust error handling.

---

## 6. Troubleshooting
- **UI freezes or Revit crashes**: You probably called the API directly from the UI thread.
- **Nothing happens on button click**: Ensure you called `Raise()` and your handler is correctly wired.
- **Need to pass data?**: Store it in the handler before calling `Raise()`.

---

## 7. References
- [Revit API Docs: ExternalEvent](https://www.revitapidocs.com/2025/6e6e2e2e-2e2e-4e2e-8e2e-2e2e2e2e2e2e.htm)
- [Jeremy Tammik: Modeless UI](https://thebuildingcoder.typepad.com/blog/2011/11/modeless-dialogs-and-external-events.html)

---

## 8. Summary
- Use `ExternalEvent` for all modeless UI → Revit API interactions.
- Keep UI and Revit logic separate.
- This pattern is safe, robust, and Revit-compliant.

---

**Teach your team: Always use ExternalEvent for UI-driven Revit API code!**
