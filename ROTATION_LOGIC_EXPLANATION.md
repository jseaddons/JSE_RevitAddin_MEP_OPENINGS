# Why MEP Element Orientation vs Wall Orientation for Rotation Logic

## The Logic Explained

### **Why Walls/Framing: MEP Element Orientation (Not Wall Orientation)**

**Answer:** We use **MEP element orientation** because we need to know **which way the MEP element is flowing** relative to the wall/framing, not which way the wall is oriented.

**Example:**
- **Wall A:** Runs horizontally (X-axis), but a duct might run vertically (Y-axis) through it
- **Wall B:** Runs vertically (Y-axis), but a duct might run horizontally (X-axis) through it

**The rotation decision depends on:**
1. **MEP Direction:** Is the duct/pipe/damper flowing X-axis or Y-axis relative to the wall?
2. **Not Wall Direction:** Whether the wall itself is X-axis or Y-axis oriented

**Code Logic:**
```csharp
// Check if MEP element is Y-axis oriented
bool isYAxisMep = Math.Abs(mepOrientation.Y) > Math.Abs(mepOrientation.X);

if (isYAxisMep)
{
    // Rotate sleeve 90° - MEP is perpendicular to wall orientation
    ElementTransformUtils.RotateElement(_doc, sleeveInstance.Id, rotationAxis, Math.PI / 2);
}
```

### **Why Floors: MEP Element Orientation**

**Answer:** Floors are **horizontal surfaces**, so MEP elements going through floors need rotation based on their **flow direction** to match the floor's coordinate system.

**Example:**
- A duct running East-West (X-axis) through a floor needs one rotation
- A duct running North-South (Y-axis) through a floor needs a different rotation

**Code Logic:**
```csharp
// FLOOR: Rotate sleeve based on MEP element direction
var mepOrientation = clashZone.MepElementOrientation;
double rotationAngle = Math.Atan2(mepOrientation.Y, mepOrientation.X);
ElementTransformUtils.RotateElement(_doc, sleeveInstance.Id, rotationAxis, rotationAngle);
```

## Summary

| Element Type | Rotation Based On | Reason |
|--------------|------------------|---------|
| **Walls/Framing** | MEP Element Orientation | Need to know if MEP is perpendicular to wall direction |
| **Floors** | MEP Element Orientation | Need to align with floor coordinate system |

**The key insight:** Rotation is about **geometric relationship** between MEP element and structural element, not the structural element's own orientation.
