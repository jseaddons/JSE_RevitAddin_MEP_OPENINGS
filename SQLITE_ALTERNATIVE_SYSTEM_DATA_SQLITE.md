# Alternative: Switch to System.Data.SQLite

> ✅ **Status**: Migration complete.  
> The project now targets `System.Data.SQLite.Core` and no longer depends on `Microsoft.Data.Sqlite`/`SQLitePCLRaw`. The notes below are kept for reference.

If `Microsoft.Data.Sqlite` continues to have issues, consider switching to `System.Data.SQLite`, which is more reliable for Revit add-ins.

## Why System.Data.SQLite?

- ✅ **More reliable** for Revit add-ins
- ✅ **Handles native libraries automatically** (no Batteries.Init() needed)
- ✅ **Better x86/x64 handling** - auto-selects correct architecture
- ✅ **Used by many production Revit add-ins**
- ✅ **Simpler setup** - fewer dependencies

## Migration Steps

### Step 1: Update NuGet Packages

In `JSE_RevitAddin_MEP_OPENINGS.csproj`:

```xml
<ItemGroup>
  <!-- Remove Microsoft.Data.Sqlite packages -->
  <!-- <PackageReference Include="Microsoft.Data.Sqlite" Version="8.0.0" /> -->
  <!-- <PackageReference Include="SQLitePCLRaw.bundle_e_sqlite3" Version="2.1.7" /> -->
  
  <!-- ✅ Use System.Data.SQLite instead -->
  <PackageReference Include="System.Data.SQLite.Core" Version="1.0.118" />
</ItemGroup>
```

### Step 2: Update Using Statements

In `Data/SleeveDbContext.cs`:

```csharp
// Change from:
using System.Data.SQLite;

// To:
using System.Data.SQLite;
```

### Step 3: Update Code

Key changes:
- `SqliteConnection` → `SQLiteConnection`
- `SqliteConnectionStringBuilder` → `SQLiteConnectionStringBuilder`
- `SqliteOpenMode` → Remove (not needed)
- `SqliteTransaction` → `SQLiteTransaction`
- Remove `Batteries.Init()` call
- Connection string format is slightly different

### Step 4: Update Connection String

```csharp
var connectionString = new SQLiteConnectionStringBuilder
{
    DataSource = _databasePath,
    Version = 3,
    JournalMode = SQLiteJournalModeEnum.Wal,
    FailIfMissing = false
}.ToString();
```

## When to Use This Alternative

Use `System.Data.SQLite` if:
- ✅ `Microsoft.Data.Sqlite` continues to fail after diagnostics
- ✅ You need a more reliable solution for production
- ✅ You want simpler deployment (auto-handles native DLLs)
- ✅ You're okay with slightly older API (but still maintained)

## Keep Microsoft.Data.Sqlite If

- ✅ You need the latest features
- ✅ You want to stay on the modern API
- ✅ Diagnostics show a fixable issue (wrong DLL, missing runtime, etc.)

## Next Steps

1. **First**: Run with enhanced diagnostics to see the actual error
2. **If error persists**: Consider switching to System.Data.SQLite
3. **If switching**: Follow migration steps above

The enhanced diagnostics will tell us exactly what's wrong, which will help decide if switching is necessary or if we can fix the current setup.

