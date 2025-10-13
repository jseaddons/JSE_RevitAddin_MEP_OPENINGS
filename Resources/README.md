# Universal Opening Families

This folder contains the 4 universal opening families required for the JSE MEP Openings add-in.

## Required Families

### 1. RectangularOpeningOnWall.rfa
- **Purpose**: Rectangular openings on walls and structural framing
- **Used for**: Rectangular ducts, rectangular pipes, cable trays, dampers on walls
- **Category**: Generic Model

### 2. CircularOpeningOnWall.rfa
- **Purpose**: Circular openings on walls and structural framing  
- **Used for**: Round ducts, round pipes on walls
- **Category**: Generic Model

### 3. RectangularOpeningOnSlab.rfa
- **Purpose**: Rectangular openings on floors/slabs
- **Used for**: Rectangular ducts, rectangular pipes, cable trays, dampers on floors
- **Category**: Generic Model

### 4. CircularOpeningOnSlab.rfa
- **Purpose**: Circular openings on floors/slabs
- **Used for**: Round ducts, round pipes on floors
- **Category**: Generic Model

## Loading Instructions

1. Open your Revit project
2. Go to **Insert > Load Family**
3. Navigate to this Resources folder
4. Load all 4 .rfa files
5. The families will appear in the Generic Model category

## Parameters

Each family includes these shared parameters:
- **Width**: Opening width (for rectangular)
- **Height**: Opening height (for rectangular)  
- **Diameter**: Opening diameter (for circular)
- **MEP_Category**: Category of MEP element (Ducts, Pipes, Cable Trays, Duct Accessories)
- **MEP_ElementId**: ID of source MEP element
- **MEP_UniqueId**: Unique ID of source MEP element
- **MEP_Size**: Formatted size (e.g., "Ø300", "400×200")
- **System_Abbreviation**: System abbreviation (HVAC, PLB, ELEC)
- **MEP_Count**: Number of MEP elements (1 for individual, >1 for clusters)
- **HostOrientation**: Orientation of host element (X, Y)

## Universal Approach Benefits

- **4 families** handle ALL MEP categories (vs 16+ category-specific families)
- **CONVOID-compatible** naming and structure
- **Future-proof** for new MEP categories (lighting, drainage, etc.)
- **Simplified scheduling** with consistent parameters
- **Reduced maintenance** - one set of families to maintain

## Deployment

When deploying the add-in:
1. Copy this Resources folder to your add-in installation directory
2. Users can load families from this folder
3. Families will be validated automatically when running the add-in


This folder contains the 4 universal opening families required for the JSE MEP Openings add-in.

## Required Families

### 1. RectangularOpeningOnWall.rfa
- **Purpose**: Rectangular openings on walls and structural framing
- **Used for**: Rectangular ducts, rectangular pipes, cable trays, dampers on walls
- **Category**: Generic Model

### 2. CircularOpeningOnWall.rfa
- **Purpose**: Circular openings on walls and structural framing  
- **Used for**: Round ducts, round pipes on walls
- **Category**: Generic Model

### 3. RectangularOpeningOnSlab.rfa
- **Purpose**: Rectangular openings on floors/slabs
- **Used for**: Rectangular ducts, rectangular pipes, cable trays, dampers on floors
- **Category**: Generic Model

### 4. CircularOpeningOnSlab.rfa
- **Purpose**: Circular openings on floors/slabs
- **Used for**: Round ducts, round pipes on floors
- **Category**: Generic Model

## Loading Instructions

1. Open your Revit project
2. Go to **Insert > Load Family**
3. Navigate to this Resources folder
4. Load all 4 .rfa files
5. The families will appear in the Generic Model category

## Parameters

Each family includes these shared parameters:
- **Width**: Opening width (for rectangular)
- **Height**: Opening height (for rectangular)  
- **Diameter**: Opening diameter (for circular)
- **MEP_Category**: Category of MEP element (Ducts, Pipes, Cable Trays, Duct Accessories)
- **MEP_ElementId**: ID of source MEP element
- **MEP_UniqueId**: Unique ID of source MEP element
- **MEP_Size**: Formatted size (e.g., "Ø300", "400×200")
- **System_Abbreviation**: System abbreviation (HVAC, PLB, ELEC)
- **MEP_Count**: Number of MEP elements (1 for individual, >1 for clusters)
- **HostOrientation**: Orientation of host element (X, Y)

## Universal Approach Benefits

- **4 families** handle ALL MEP categories (vs 16+ category-specific families)
- **CONVOID-compatible** naming and structure
- **Future-proof** for new MEP categories (lighting, drainage, etc.)
- **Simplified scheduling** with consistent parameters
- **Reduced maintenance** - one set of families to maintain

## Deployment

When deploying the add-in:
1. Copy this Resources folder to your add-in installation directory
2. Users can load families from this folder
3. Families will be validated automatically when running the add-in
















