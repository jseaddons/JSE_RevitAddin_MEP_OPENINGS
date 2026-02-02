# Revit Multi-Version Integration Guide (2023-2026)

This guide explains how to set up a Revit Add-in project to support multiple Revit versions (2023, 2024, 2025, and 2026) using modern .NET SDKs and the Nice3point extension ecosystem.

## 1. Project Skeleton (.csproj)

Use a **Microsoft.NET.Sdk** project. Key settings:

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <Configurations>Debug R23;Debug R24;Debug R25;Debug R26;Release R23;Release R24;Release R25;Release R26</Configurations>
    <ImplicitUsings>true</ImplicitUsings>
    <LangVersion>latest</LangVersion>
    <!-- UseWPF/WindowsForms as needed -->
    <UseWPF>true</UseWPF> 
    <UseWindowsForms>true</UseWindowsForms>
  </PropertyGroup>
```

## 2. Framework Targeting

Each Revit version requires a specific .NET version:

| Revit Version | .NET Version | TargetFramework |
| :--- | :--- | :--- |
| **2023** | .NET Framework 4.8 | `net48` |
| **2024** | .NET 6.0 | `net6.0-windows` |
| **2025** | .NET 8.0 | `net8.0-windows` |
| **2026** | .NET 8.0 | `net8.0-windows` |

**Implementation in .csproj:**

```xml
<PropertyGroup Condition="$(Configuration.Contains('R23'))">
  <RevitVersion>2023</RevitVersion>
  <TargetFramework>net48</TargetFramework>
  <DefineConstants>$(DefineConstants);REVIT2023</DefineConstants>
</PropertyGroup>

<PropertyGroup Condition="$(Configuration.Contains('R24'))">
  <RevitVersion>2024</RevitVersion>
  <TargetFramework>net6.0-windows</TargetFramework>
  <DefineConstants>$(DefineConstants);REVIT2024</DefineConstants>
</PropertyGroup>
<!-- Repeat for R25/R26 with net8.0-windows -->
```

## 3. Nice3point Package Management

The "Annointing" trick is to group packages by configuration so the correct Revit API version is used:

```xml
<ItemGroup Condition="$(Configuration.Contains('R23'))">
  <PackageReference Include="Nice3point.Revit.Toolkit" Version="2023.*" />
  <PackageReference Include="Nice3point.Revit.Extensions" Version="2023.*" />
</ItemGroup>

<ItemGroup Condition="$(Configuration.Contains('R24'))">
  <PackageReference Include="Nice3point.Revit.Toolkit" Version="2024.*" />
  <PackageReference Include="Nice3point.Revit.Extensions" Version="2024.*" />
</ItemGroup>
```

## 4. Resolving Conflicts (The "TaskDialog" Fix)

In .NET 6+ (Revit 2024+), `System.Windows.Forms` includes a `TaskDialog` which conflicts with `Autodesk.Revit.UI.TaskDialog`. Fix this globally in the `.csproj`:

```xml
<ItemGroup Condition="'$(TargetFramework)' != 'net48'">
  <Using Include="Autodesk.Revit.UI.TaskDialog" Alias="TaskDialog" />
  <Using Include="Autodesk.Revit.UI.TaskDialogResult" Alias="TaskDialogResult" />
  <Using Include="Autodesk.Revit.UI.TaskDialogCommonButtons" Alias="TaskDialogCommonButtons" />
  <Using Include="Autodesk.Revit.UI.TaskDialogIcon" Alias="TaskDialogIcon" />
</ItemGroup>
```

## 5. Version-Specific Code

Use the preprocessor directives defined in step 2:

```csharp
#if REVIT2023
    // Old API logic
#else
    // Modern API logic (2024+)
#endif
```

## 6. Post-Build Deployment

Automate the copy to the Revit Addins folder:

```xml
<Target Name="PostBuild" AfterTargets="Build">
  <PropertyGroup>
    <RevitAddinFolder>C:\ProgramData\Autodesk\Revit\Addins\$(RevitVersion)\</RevitAddinFolder>
  </PropertyGroup>
  <Copy SourceFiles="$(OutputPath)$(AssemblyName).dll" DestinationFolder="$(RevitAddinFolder)" />
  <Copy SourceFiles="$(MSBuildProjectDirectory)\$(AssemblyName).addin" DestinationFolder="$(RevitAddinFolder)" />
</Target>
```
