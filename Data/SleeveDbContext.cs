using System;
using System.Collections.Generic;
#if !NET8_0_OR_GREATER
using System.Data.SQLite;
#endif
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Data.Entities;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Data
{
    /// <summary>
    /// SQLite database context for sleeve persistence
    /// Phase SQLite-1: Dual-write mode (writes to both XML and SQLite)
    /// </summary>
    public class SleeveDbContext : IDisposable
    {
        private readonly SQLiteConnection _connection;
        private readonly string _databasePath;
        private readonly Action<string> _logger;
        private bool _disposed;
        

        /// <summary>
        /// ✅ IST TIMEZONE: Convert UTC to IST (Indian Standard Time = UTC+5:30)
        /// </summary>
        private static DateTime GetIstTime()
        {
            return DateTime.UtcNow.AddHours(5.5);
        }

        /// <summary>
        /// ✅ IST TIMEZONE: Get IST timestamp string for SQLite (format: 'YYYY-MM-DD HH:MM:SS')
        /// </summary>
        private static string GetIstTimestamp()
        {
            return GetIstTime().ToString("yyyy-MM-dd HH:mm:ss");
        }

        /// <summary>
        /// Initialize database context for a Revit document
        /// Database file: {ProjectName}_SleevePersistence.db in Filters directory
        /// </summary>
        public SleeveDbContext(Document document, Action<string>? logger = null)
        {
            if (document == null) throw new ArgumentNullException(nameof(document));

            _logger = logger ?? (_ => { });

            try
            {
                var assemblyLocation = Assembly.GetExecutingAssembly().Location;
                var assemblyDirectory = Path.GetDirectoryName(assemblyLocation) ?? string.Empty;

                _logger($"[SQLite] Assembly directory: {assemblyDirectory}");

#if NET8_0_OR_GREATER
                // ✅ CRITICAL: Initialize SQLitePCL for .NET 8 (Revit 2025+)
                // This resolves the "You need to call SQLitePCL.raw.SetProvider()" error in Revit plugins.
                try { SQLitePCL.Batteries_V2.Init(); } catch (Exception pex) { _logger($"[SQLite] ⚠️ SQLitePCL Init Warning: {pex.Message}"); }
#endif

#if !NET8_0_OR_GREATER
                // System.Data.SQLite native DLL verification (not needed for Microsoft.Data.Sqlite on NET8+)
                VerifyDependency("System.Data.SQLite.dll", assemblyDirectory);
                VerifyDependency(Path.Combine("x64", "SQLite.Interop.dll"), assemblyDirectory);
#endif

                var filtersDirectory = ProjectPathService.GetFiltersDirectory(document);
                Directory.CreateDirectory(filtersDirectory);
                _logger($"[SQLite] Filters directory: {filtersDirectory}");

                var projectName = document.Title ?? "Default";
                var safeProjectName = Regex.Replace(projectName, @"[^\w\s-]", string.Empty).Replace(" ", "_");
                _databasePath = Path.Combine(filtersDirectory, $"{safeProjectName}_SleevePersistence.db");
                _logger($"[SQLite] Database path: {_databasePath}");

#if NET8_0_OR_GREATER
                // Microsoft.Data.Sqlite: plain connection string; ForeignKeys/JournalMode set via PRAGMAs below
                _connection = new SQLiteConnection($"Data Source={_databasePath}");
#else
                var builder = new SQLiteConnectionStringBuilder
                {
                    DataSource = _databasePath,
                    ForeignKeys = true,
                    JournalMode = SQLiteJournalModeEnum.Wal
                };
                _connection = new SQLiteConnection(builder.ConnectionString);
#endif
                _connection.Open();
                _logger("[SQLite] ✅ Connection opened successfully");

                // ✅ PERFORMANCE: Configure SQLite for maximum performance
                using (var cmd = _connection.CreateCommand())
                {
                    cmd.CommandText = "PRAGMA foreign_keys = ON;";
                    cmd.ExecuteNonQuery();
                }

                using (var cmd = _connection.CreateCommand())
                {
                    cmd.CommandText = "PRAGMA journal_mode = WAL;";
                    cmd.ExecuteNonQuery();
                }

                // 🚀 PERFORMANCE BOOST: Increase cache size to 64MB (default is 2MB)
                // Larger cache = fewer disk I/O operations, especially for large datasets
                using (var cmd = _connection.CreateCommand())
                {
                    cmd.CommandText = "PRAGMA cache_size = -65536;"; // Negative = KB, 64MB cache
                    cmd.ExecuteNonQuery();
                }

                // 🚀 PERFORMANCE BOOST: Use NORMAL synchronous mode instead of FULL
                // WAL mode already provides crash safety, NORMAL is 2-3x faster
                using (var cmd = _connection.CreateCommand())
                {
                    cmd.CommandText = "PRAGMA synchronous = NORMAL;";
                    cmd.ExecuteNonQuery();
                }

                // 🚀 PERFORMANCE BOOST: Use memory for temp storage instead of disk
                // Speeds up complex queries with sorting/grouping
                using (var cmd = _connection.CreateCommand())
                {
                    cmd.CommandText = "PRAGMA temp_store = MEMORY;";
                    cmd.ExecuteNonQuery();
                }

                // 🚀 PERFORMANCE BOOST: Increase page size to 8KB (default 4KB)
                // Better for larger records like ClashZones with many fields
                // NOTE: This only works on new databases, existing DBs keep their page size
                using (var cmd = _connection.CreateCommand())
                {
                    cmd.CommandText = "PRAGMA page_size = 8192;";
                    cmd.ExecuteNonQuery();
                }

                // 🚀 PERFORMANCE BOOST: Set mmap_size to 256MB for memory-mapped I/O
                // Significantly faster reads, especially for queries
                using (var cmd = _connection.CreateCommand())
                {
                    cmd.CommandText = "PRAGMA mmap_size = 268435456;"; // 256MB
                    cmd.ExecuteNonQuery();
                }

                // ⚡ OPTIMIZATION: Verify schema on context creation
                // We use IF NOT EXISTS and ColumnExists checks which are efficient.
                EnsureSchemaCreated();
                EnsureSchemaUpgraded();
                CheckRTreeSupport();

                if (File.Exists(_databasePath))
                {
                    var info = new FileInfo(_databasePath);
                    _logger($"[SQLite] ✅ Database ready ({info.Length} bytes)");
                }
            }
            catch (Exception ex)
            {
                _logger($"[SQLite] ❌ Error during initialization: {ex.Message}");
                _logger($"[SQLite] Stack trace: {ex.StackTrace}");
                
                // Show user-friendly error before throwing
                var userMessage = $"Database connection failed:\n\n{ex.Message}\n\n" +
                    "This can happen when:\n" +
                    "• SQLite DLLs are missing (x64\\SQLite.Interop.dll)\n" +
                    "• Database file is locked by another Revit instance\n" +
                    "• Antivirus is blocking file access\n\n" +
                    "Check the log file for details.";
                    
                System.Windows.Forms.MessageBox.Show(
                    userMessage,
                    "❌ Database Connection Failed",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
                
                throw;
            }
        }

        /// <summary>
        /// Lightweight constructor for background-thread usage.
        /// Accepts a pre-resolved database path (no Revit Document dependency).
        /// Skips schema verification (assumes DB already exists and is migrated).
        /// </summary>
        public SleeveDbContext(string databasePath, Action<string>? logger = null)
        {
            if (string.IsNullOrEmpty(databasePath)) throw new ArgumentNullException(nameof(databasePath));
            _logger = logger ?? (_ => { });
            _databasePath = databasePath;

#if NET8_0_OR_GREATER
            // ✅ CRITICAL: Initialize SQLitePCL for .NET 8 (Revit 2025+)
            try { SQLitePCL.Batteries_V2.Init(); } catch (Exception pex) { _logger($"[SQLite] ⚠️ SQLitePCL Init Warning: {pex.Message}"); }
            _connection = new SQLiteConnection($"Data Source={_databasePath}");
#else
            var builder = new SQLiteConnectionStringBuilder
            {
                DataSource = _databasePath,
                ForeignKeys = true,
                JournalMode = SQLiteJournalModeEnum.Wal
            };
            _connection = new SQLiteConnection(builder.ConnectionString);
#endif
            _connection.Open();

            using (var cmd = _connection.CreateCommand()) { cmd.CommandText = "PRAGMA foreign_keys = ON;"; cmd.ExecuteNonQuery(); }
            using (var cmd = _connection.CreateCommand()) { cmd.CommandText = "PRAGMA journal_mode = WAL;"; cmd.ExecuteNonQuery(); }
            using (var cmd = _connection.CreateCommand()) { cmd.CommandText = "PRAGMA synchronous = NORMAL;"; cmd.ExecuteNonQuery(); }
            using (var cmd = _connection.CreateCommand()) { cmd.CommandText = "PRAGMA cache_size = -65536;"; cmd.ExecuteNonQuery(); }
            using (var cmd = _connection.CreateCommand()) { cmd.CommandText = "PRAGMA temp_store = MEMORY;"; cmd.ExecuteNonQuery(); }
            using (var cmd = _connection.CreateCommand()) { cmd.CommandText = "PRAGMA mmap_size = 268435456;"; cmd.ExecuteNonQuery(); }
        }

        private void VerifyDependency(string relativePath, string assemblyDirectory)
        {
            var targetPath = Path.Combine(assemblyDirectory, relativePath);
            if (File.Exists(targetPath))
            {
                _logger($"[SQLite] ✅ Found dependency: {relativePath}");
                // ✅ Explicitly load native library to satisfy SQLite interop
                if (relativePath.EndsWith("SQLite.Interop.dll", StringComparison.OrdinalIgnoreCase))
                {
                    Helpers.NativeLibraryLoader.LoadNativeLibrary(targetPath);
                }
                return;
            }

            _logger($"[SQLite] ⚠️ Missing dependency: {relativePath}");
            var candidates = FindInNuGet(relativePath);
            foreach (var candidate in candidates)
            {
                try
                {
                    var destination = Path.Combine(assemblyDirectory, relativePath);
                    Directory.CreateDirectory(Path.GetDirectoryName(destination) ?? assemblyDirectory);
                    File.Copy(candidate, destination, overwrite: true);
                    _logger($"[SQLite] ✅ Copied dependency from NuGet cache: {candidate}");
                    
                    // ✅ Explicitly load native library after copy
                    if (relativePath.EndsWith("SQLite.Interop.dll", StringComparison.OrdinalIgnoreCase))
                    {
                        Helpers.NativeLibraryLoader.LoadNativeLibrary(destination);
                    }
                    return;
                }
                catch (Exception copyEx)
                {
                    _logger($"[SQLite] ❌ Failed to copy dependency from {candidate}: {copyEx.Message}");
                }
            }

            _logger($"[SQLite] ⚠️ Dependency {relativePath} remains missing. SQLite may fail to load.");
        }

        private static IEnumerable<string> FindInNuGet(string relativePath)
        {
            var fileName = Path.GetFileName(relativePath);
             var nugetRoot = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            if (string.IsNullOrEmpty(nugetRoot))
            {
                yield break;
            }

            var sqlitePackageRoot = Path.Combine(nugetRoot, @".nuget\packages\system.data.sqlite.core");
            // Check if root exists
            if (!Directory.Exists(sqlitePackageRoot))
            {
                yield break;
            }

            IEnumerable<string> files = null;
            try 
            {
                // Fix: Search for the filename only, not the full relative path (e.g. "x64\SQLite.Interop.dll" -> "SQLite.Interop.dll")
                // EnumerateFiles throws if the pattern contains a path to a non-existent directory.
                files = Directory.EnumerateFiles(sqlitePackageRoot, fileName, SearchOption.AllDirectories);
            }
            catch 
            {
                 yield break; 
            }

            foreach (var candidate in files)
            {
                // Special handling for Interop to ensure we pick x64 or x86 correctly if needed
                if (fileName.Equals("SQLite.Interop.dll", StringComparison.OrdinalIgnoreCase))
                {
                    // Strict check: if looking for x64, ensure path contains x64
                    // (Assuming User wants 64-bit for Revit)
                     if (!candidate.Contains("x64", StringComparison.OrdinalIgnoreCase)) continue;
                }

                yield return candidate;
            }
        }

        /// <summary>
        /// Ensure database schema is created
        /// </summary>
        private void EnsureSchemaCreated()
        {
            using (var transaction = _connection.BeginTransaction())
            {
                try
                {
                    // Create Filters table
                    ExecuteCommand(@"
                        CREATE TABLE IF NOT EXISTS Filters (
                            FilterId          INTEGER PRIMARY KEY AUTOINCREMENT,
                            FilterName        TEXT NOT NULL,
                            Category          TEXT NOT NULL,
                            ReferenceDocKey   TEXT,
                            HostDocKey        TEXT,
                            ReferenceCategory TEXT,
                            SelectedHostCategories TEXT,
                            AdoptToDocumentFlag INTEGER DEFAULT 1,
                            IsFilterComboNew  INTEGER NOT NULL DEFAULT 1,
                            CreatedAt         DATETIME NOT NULL DEFAULT (datetime('now', '+5 hours', '+30 minutes')),
                            UpdatedAt         DATETIME NOT NULL DEFAULT (datetime('now', '+5 hours', '+30 minutes')),
                            UNIQUE(FilterName, Category)
                        )", transaction);

                    // Create FileCombos table
                    ExecuteCommand(@"
                        CREATE TABLE IF NOT EXISTS FileCombos (
                            ComboId              INTEGER PRIMARY KEY AUTOINCREMENT,
                            FilterId             INTEGER NOT NULL,
                            Category             TEXT NOT NULL,
                            SelectedHostCategories TEXT,
                            LinkedFileKey        TEXT NOT NULL,
                            HostFileKey          TEXT NOT NULL,
                            IsFilterComboNew     INTEGER NOT NULL DEFAULT 1,
                            ProcessedAt          DATETIME,
                            CreatedAt            DATETIME DEFAULT CURRENT_TIMESTAMP,
                            UpdatedAt            DATETIME DEFAULT CURRENT_TIMESTAMP,
                            FOREIGN KEY(FilterId) REFERENCES Filters(FilterId) ON DELETE CASCADE,
                            UNIQUE(FilterId, Category, LinkedFileKey, HostFileKey)
                        )", transaction);
                    
                    // ✅ MIGRATION: Add Category and SelectedHostCategories columns if they don't exist (for existing databases)
                    try
                    {
                        ExecuteCommand(@"
                            ALTER TABLE FileCombos ADD COLUMN Category TEXT;
                        ", transaction);
                    }
                    catch
                    {
                        // Column already exists, ignore
                    }
                    
                    try
                    {
                        ExecuteCommand(@"
                            ALTER TABLE FileCombos ADD COLUMN SelectedHostCategories TEXT;
                        ", transaction);
                    }
                    catch
                    {
                        // Column already exists, ignore
                    }

                    // Create ClashZones table
                    ExecuteCommand(@"
                        CREATE TABLE IF NOT EXISTS ClashZones (
                            ClashZoneId   INTEGER PRIMARY KEY AUTOINCREMENT,
                            ComboId       INTEGER NOT NULL,
                            ClashZoneGuid TEXT NOT NULL DEFAULT '',
                            MepElementId  INTEGER NOT NULL,
                            HostElementId INTEGER NOT NULL,
                            IntersectionX REAL NOT NULL,
                            IntersectionY REAL NOT NULL,
                            IntersectionZ REAL NOT NULL,
                            SleeveState   INTEGER NOT NULL DEFAULT 0,
                            SleeveInstanceId INTEGER DEFAULT -1,
                            ClusterInstanceId INTEGER DEFAULT -1,
                            SleeveWidth   REAL,
                            SleeveHeight  REAL,
                            SleeveDiameter REAL,
                            SleeveDepth REAL,
                            SleevePlacementX REAL,
                            SleevePlacementY REAL,
                            SleevePlacementZ REAL,
                            SleevePlacementActiveX REAL,
                            SleevePlacementActiveY REAL,
                            SleevePlacementActiveZ REAL,
                            BoundingBoxMinX REAL,
                            BoundingBoxMinY REAL,
                            BoundingBoxMinZ REAL,
                            BoundingBoxMaxX REAL,
                            BoundingBoxMaxY REAL,
                            BoundingBoxMaxZ REAL,
                            RotatedBoundingBoxMinX REAL,
                            RotatedBoundingBoxMinY REAL,
                            RotatedBoundingBoxMinZ REAL,
                            RotatedBoundingBoxMaxX REAL,
                            RotatedBoundingBoxMaxY REAL,
                            RotatedBoundingBoxMaxZ REAL,
                            MepRotationCos REAL,
                            MepRotationSin REAL,
                            SleeveCorner1X REAL,
                            SleeveCorner1Y REAL,
                            SleeveCorner1Z REAL,
                            SleeveCorner2X REAL,
                            SleeveCorner2Y REAL,
                            SleeveCorner2Z REAL,
                            SleeveCorner3X REAL,
                            SleeveCorner3Y REAL,
                            SleeveCorner3Z REAL,
                            SleeveCorner4X REAL,
                            SleeveCorner4Y REAL,
                            SleeveCorner4Z REAL,
                            PlacementSource TEXT,
                            StructuralThickness REAL DEFAULT 0.0,
                            WallThickness REAL DEFAULT 0.0,
                            FramingThickness REAL DEFAULT 0.0,
                            MepSystemType TEXT,
                            MepSystemName TEXT,
                            MepServiceType TEXT,
                            ElevationFromLevel REAL DEFAULT 0.0,
                            MepCategory TEXT,
                            StructuralType TEXT,
                            HostOrientation TEXT,
                            WallDirectionType TEXT,
                            MepOrientationDirection TEXT,
                            MepOrientationX REAL,
                            MepOrientationY REAL,
                            MepOrientationZ REAL,
                            MepRotationAngleRad REAL,
                            MepRotationAngleDeg REAL,
                            MepAngleToXRad REAL,
                            MepAngleToXDeg REAL,
                            MepAngleToYRad REAL,
                            MepAngleToYDeg REAL,
                            MepWidth REAL,
                            MepHeight REAL,
                            MepElementLevelName TEXT,
                            MepElementLevelElevation REAL DEFAULT 0.0,
                            WallCenterlinePointX REAL DEFAULT 0.0,
                            WallCenterlinePointY REAL DEFAULT 0.0,
                            WallCenterlinePointZ REAL DEFAULT 0.0,
                            PlacementStatus TEXT DEFAULT 'NotReady',
                            CalculatedSleeveWidth REAL,
                            CalculatedSleeveHeight REAL,
                            CalculatedSleeveDiameter REAL,
                            CalculatedSleeveDepth REAL,
                            CalculatedRotation REAL,
                            CalculatedPlacementX REAL,
                            CalculatedPlacementY REAL,
                            CalculatedPlacementZ REAL,
                            CalculatedFamilyName TEXT,
                            ValidationStatus TEXT DEFAULT 'Valid',
                            ValidationMessage TEXT,
                            CalculationBatchId TEXT,
                            CalculatedAt TEXT,
                            PlacedAt TEXT,
                            IsCurrentClashFlag INTEGER NOT NULL DEFAULT 0,
                            ReadyForPlacementFlag INTEGER NOT NULL DEFAULT 0,
                            IsResolvedFlag INTEGER NOT NULL DEFAULT 0,
                            IsClusterResolvedFlag INTEGER NOT NULL DEFAULT 0,
                            IsCombinedResolved INTEGER NOT NULL DEFAULT 0,
                            IsClusteredFlag INTEGER NOT NULL DEFAULT 0,
                            MarkedForClusterProcess INTEGER,
                            AfterClusterSleeveId INTEGER NOT NULL DEFAULT -1,
                            CombinedClusterSleeveInstanceId INTEGER,
                            HasDamperNearbyFlag INTEGER NOT NULL DEFAULT 0,
                            HasMepConnector INTEGER DEFAULT 0,
                            DamperConnectorSide TEXT DEFAULT '',
                            IsInsulated INTEGER DEFAULT 0,
                            InsulationThickness REAL DEFAULT 0.0,
                            MepParameterValuesJson TEXT,
                            HostParameterValuesJson TEXT,
                            MepElementTypeName TEXT,
                            MepElementFamilyName TEXT,
                            MepElementSizeParameterValue TEXT DEFAULT '',
                            SourceDocKey TEXT,
                            HostDocKey TEXT,
                            MepElementUniqueId TEXT,
                            UpdatedAt     DATETIME NOT NULL DEFAULT (datetime('now', '+5 hours', '+30 minutes')),
                            FOREIGN KEY(ComboId) REFERENCES FileCombos(ComboId) ON DELETE CASCADE,
                            UNIQUE(ComboId, MepElementId, HostElementId, IntersectionX, IntersectionY, IntersectionZ)
                        )", transaction);

                    // Create indexes for performance
                    ExecuteCommand("CREATE INDEX IF NOT EXISTS idx_clashzones_sleevestate ON ClashZones(SleeveState)", transaction);
                    ExecuteCommand("CREATE INDEX IF NOT EXISTS idx_clashzones_mep_element ON ClashZones(MepElementId)", transaction);
                    ExecuteCommand("CREATE INDEX IF NOT EXISTS idx_clashzones_host_element ON ClashZones(HostElementId)", transaction);
                    ExecuteCommand("CREATE INDEX IF NOT EXISTS idx_clashzones_sleeve_instance ON ClashZones(SleeveInstanceId)", transaction);
                    ExecuteCommand("CREATE INDEX IF NOT EXISTS idx_clashzones_cluster_instance ON ClashZones(ClusterInstanceId)", transaction);
                    ExecuteCommand("CREATE INDEX IF NOT EXISTS idx_clashzones_mep_host_point ON ClashZones(MepElementId, HostElementId, IntersectionX, IntersectionY, IntersectionZ)", transaction);
                    
                    // ✅ NOTE: GUID indexes are created in EnsureSchemaUpgraded() after ClashZoneGuid column is added

                    // Create SleeveEvents table
                    ExecuteCommand(@"
                        CREATE TABLE IF NOT EXISTS SleeveEvents (
                            EventId       INTEGER PRIMARY KEY AUTOINCREMENT,
                            ClashZoneId   INTEGER NOT NULL,
                            EventType     TEXT NOT NULL,
                            Payload       TEXT,
                            CreatedAt     DATETIME NOT NULL DEFAULT (datetime('now', '+5 hours', '+30 minutes')),
                            FOREIGN KEY(ClashZoneId) REFERENCES ClashZones(ClashZoneId) ON DELETE CASCADE
                        )", transaction);

                    // Create Conditions table
                    ExecuteCommand(@"
                        CREATE TABLE IF NOT EXISTS Conditions (
                            ConditionId                       INTEGER PRIMARY KEY AUTOINCREMENT,
                            FilterId                          INTEGER NOT NULL,
                            Category                          TEXT NOT NULL,
                            RectNormal                        REAL,
                            RectInsulated                     REAL,
                            RoundNormal                       REAL,
                            RoundInsulated                    REAL,
                            PipesNormal                       REAL,
                            PipesInsulated                    REAL,
                            UseNominalDiameterForPipes        INTEGER,
                            CableTrayTop                      REAL,
                            CableTrayTopInsulated             REAL,
                            CableTrayOther                    REAL,
                            CableTrayOtherInsulated           REAL,
                            DuctAccessoryMepNormal            REAL,
                            DuctAccessoryMepInsulated         REAL,
                            DuctAccessoryOtherNormal          REAL,
                            DuctAccessoryOtherInsulated       REAL,
                            OpeningPrefs                      TEXT,
                            -- Per-filter/category behavioral settings
                            JoinOpeningsDistanceMm            REAL,
                            IgnoreArchitecturalFloors         INTEGER,
                            CircularToRectangularThresholdMm  REAL,
                            UpdatedAt                         DATETIME NOT NULL DEFAULT (datetime('now', '+5 hours', '+30 minutes')),
                            FOREIGN KEY(FilterId) REFERENCES Filters(FilterId) ON DELETE CASCADE,
                            UNIQUE(FilterId, Category)
                        )", transaction);

                    EnsureSleeveSnapshotTable(transaction);
                    EnsureParameterTransferFlagsTable(transaction);
                    EnsureCategoryProcessingMarkersTable(transaction);
                    EnsureClusterSleevesTable(transaction);
                    EnsureClusterSleevesV2Table(transaction); // ✅ BATCH V2: Cluster calculation table
                    EnsureCombinedSleevesTables(transaction);
                    EnsureSessionContextTable(transaction); // ✅ SESSION CONTEXT: Store section box bounds and session data
                    
                    // ✅ R-TREE: Create R-tree virtual table for spatial indexing (if enabled)
                    EnsureRTreeTable(transaction);

                    // Create migration tracking table
                    ExecuteCommand(@"
                        CREATE TABLE IF NOT EXISTS SchemaMigrations (
                            MigrationId   INTEGER PRIMARY KEY AUTOINCREMENT,
                            Version       TEXT NOT NULL UNIQUE,
                            AppliedAt     DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
                        )", transaction);

                    transaction.Commit();
                    _logger($"[SQLite] ✅ Schema created/verified for database: {_databasePath}");
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    _logger($"[SQLite] ❌ Error creating schema: {ex.Message}");
                    throw;
                }
            }
        }

        private void ExecuteCommand(string sql, SQLiteTransaction transaction)
        {
            using (var cmd = _connection.CreateCommand())
            {
                cmd.CommandText = sql;
                cmd.Transaction = transaction;
                cmd.ExecuteNonQuery();
            }
        }

        /// <summary>
        /// ✅ R-TREE CHECK: Test if SQLite R-tree extension is available and supported
        /// R-tree is built into SQLite but may need to be enabled in some builds
        /// </summary>
        private void CheckRTreeSupport()
        {
            try
            {
                // Step 1: Check SQLite version (R-tree has been available since SQLite 3.5.0)
                string sqliteVersion = string.Empty;
                using (var cmd = _connection.CreateCommand())
                {
                    cmd.CommandText = "SELECT sqlite_version()";
                    var result = cmd.ExecuteScalar();
                    sqliteVersion = result?.ToString() ?? "unknown";
                }
                _logger($"[SQLite] SQLite version: {sqliteVersion}");

                // Step 2: Try to create a test R-tree virtual table
                // If this succeeds, R-tree is available
                bool rtreeSupported = false;
                string errorMessage = string.Empty;
                
                try
                {
                    using (var cmd = _connection.CreateCommand())
                    {
                        // Create a temporary test R-tree table
                        cmd.CommandText = @"
                            CREATE VIRTUAL TABLE IF NOT EXISTS _rtree_test USING rtree(
                                id INTEGER PRIMARY KEY,
                                minX REAL, maxX REAL,
                                minY REAL, maxY REAL,
                                minZ REAL, maxZ REAL
                            )";
                        cmd.ExecuteNonQuery();
                        
                        // Try to insert a test record
                        cmd.CommandText = @"
                            INSERT INTO _rtree_test (id, minX, maxX, minY, maxY, minZ, maxZ)
                            VALUES (1, 0.0, 1.0, 0.0, 1.0, 0.0, 1.0)";
                        cmd.ExecuteNonQuery();
                        
                        // Try to query it
                        cmd.CommandText = @"
                            SELECT COUNT(*) FROM _rtree_test 
                            WHERE minX <= 1.0 AND maxX >= 0.0
                              AND minY <= 1.0 AND maxY >= 0.0
                              AND minZ <= 1.0 AND maxZ >= 0.0";
                        var count = cmd.ExecuteScalar();
                        
                        // Clean up test table
                        cmd.CommandText = "DROP TABLE IF EXISTS _rtree_test";
                        cmd.ExecuteNonQuery();
                        
                        rtreeSupported = true;
                        _logger($"[SQLite] ✅ R-tree extension is SUPPORTED and working correctly");
                    }
                }
                catch (Exception ex)
                {
                    rtreeSupported = false;
                    errorMessage = ex.Message;
                    
                    // Clean up test table if it was partially created
                    try
                    {
                        using (var cmd = _connection.CreateCommand())
                        {
                            cmd.CommandText = "DROP TABLE IF EXISTS _rtree_test";
                            cmd.ExecuteNonQuery();
                        }
                    }
                    catch { }
                    
                    _logger($"[SQLite] ❌ R-tree extension is NOT available: {errorMessage}");
                }

                // Step 3: Log detailed information
                if (rtreeSupported)
                {
                    _logger($"[SQLite] ✅ R-tree is ready for spatial indexing implementation");
                    _logger($"[SQLite]    - Can create R-tree virtual tables");
                    _logger($"[SQLite]    - Can perform spatial queries (bounding box, proximity)");
                    _logger($"[SQLite]    - Recommended for section box filtering optimization");
                }
                else
                {
                    _logger($"[SQLite] ⚠️ R-tree is NOT available - spatial queries will use in-memory filtering");
                    _logger($"[SQLite]    - Section box filtering will load all zones, then filter in C#");
                    _logger($"[SQLite]    - Consider upgrading SQLite if R-tree support is needed");
                }
            }
            catch (Exception ex)
            {
                _logger($"[SQLite] ⚠️ Error checking R-tree support: {ex.Message}");
                _logger($"[SQLite]    Assuming R-tree is NOT available (fallback to in-memory filtering)");
            }
        }

        private void EnsureSchemaUpgraded()
        {
            try
            {
                using (var transaction = _connection.BeginTransaction())
                {
                    EnsureSleeveSnapshotTable(transaction);
                    EnsureParameterTransferFlagsTable(transaction);
                    EnsureCategoryProcessingMarkersTable(transaction);
                    EnsureClusterSleevesTable(transaction);
                    EnsureClusterSleevesV2Table(transaction); // ✅ BATCH V2: New table for decoupled calculation
                    
                    // ✅ CRITICAL: Call EnsureCombinedSleevesTables BEFORE any AddColumnIfMissing calls
                    // This ensures the CombinedSleeves table exists before any code tries to modify it
                    // Fixes "no such table: CombinedSleeves" error on fresh installations
                    EnsureCombinedSleevesTables(transaction);
                    
                    // ✅ R-TREE: Create R-tree table for existing databases (if enabled)
                    EnsureRTreeTable(transaction);

                    // ✅ TRIGGERS: Ensure flag management triggers are up to date (drops stale triggers)
                    EnsureFlagManagementTriggers(transaction);

                    AddColumnIfMissing("ClashZones", "ClashZoneGuid", "TEXT NOT NULL DEFAULT ''", transaction);
                    AddColumnIfMissing("ClashZones", "MepCategory", "TEXT", transaction);
                    AddColumnIfMissing("ClashZones", "StructuralType", "TEXT", transaction);
                    AddColumnIfMissing("ClashZones", "HostOrientation", "TEXT", transaction);
                    AddColumnIfMissing("ClashZones", "MepOrientationDirection", "TEXT", transaction);
                    AddColumnIfMissing("ClashZones", "MepOrientationX", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "MepOrientationY", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "MepOrientationZ", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "MepRotationAngleRad", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "MepRotationAngleDeg", "REAL", transaction);
                    
                    // ✅ UNIFIED BATCH MODE (PHASE 3): New columns for individual sleeve calculation & persistence
                    if (AddColumnIfMissing("ClashZones", "PlacementStatus", "TEXT DEFAULT 'NotReady'", transaction))
                        _logger("[SQLite] ✅ Added PlacementStatus column to ClashZones");
                        
                    AddColumnIfMissing("ClashZones", "CalculatedSleeveWidth", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "CalculatedSleeveHeight", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "CalculatedSleeveDiameter", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "CalculatedSleeveDepth", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "SleeveDepth", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "CalculatedRotation", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "CalculatedPlacementX", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "CalculatedPlacementY", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "CalculatedPlacementZ", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "CalculatedFamilyName", "TEXT", transaction);
                    AddColumnIfMissing("ClashZones", "ValidationStatus", "TEXT DEFAULT 'Valid'", transaction);
                    AddColumnIfMissing("ClashZones", "ValidationMessage", "TEXT", transaction);
                    AddColumnIfMissing("ClashZones", "CalculationBatchId", "TEXT", transaction);
                    AddColumnIfMissing("ClashZones", "CalculatedAt", "TEXT", transaction);
                    AddColumnIfMissing("ClashZones", "PlacedAt", "TEXT", transaction);

                    // ✅ ROTATION MATRIX: Pre-calculated cos/sin for "dump once use many times" principle
                    // Calculated once during placement, stored for reuse during clustering (avoids repeated Math.Cos/Sin calls)
                    AddColumnIfMissing("ClashZones", "MepRotationCos", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "MepRotationSin", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "MepAngleToXRad", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "MepAngleToXDeg", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "MepAngleToYRad", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "MepAngleToYDeg", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "MepWidth", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "MepHeight", "REAL", transaction);
                    // ✅ REFERENCE LEVEL: Add MEP element Reference Level column
                    AddColumnIfMissing("ClashZones", "MepElementLevelName", "TEXT", transaction);
                    // ✅ REFERENCE LEVEL ELEVATION: Add MEP element Reference Level elevation (critical for Elevation from Level and Bottom of Opening calculation)
                    if (AddColumnIfMissing("ClashZones", "MepElementLevelElevation", "REAL DEFAULT 0.0", transaction))
                        _logger("[SQLite] ✅ Added MepElementLevelElevation column to ClashZones (Reference Level elevation in feet)");
                    // ✅ WALL CENTERLINE POINT: Pre-calculated during refresh to enable multi-threaded placement
                    if (AddColumnIfMissing("ClashZones", "WallCenterlinePointX", "REAL DEFAULT 0.0", transaction))
                        _logger("[SQLite] ✅ Added WallCenterlinePointX column to ClashZones (pre-calculated wall centerline point)");
                    if (AddColumnIfMissing("ClashZones", "WallCenterlinePointY", "REAL DEFAULT 0.0", transaction))
                        _logger("[SQLite] ✅ Added WallCenterlinePointY column to ClashZones (pre-calculated wall centerline point)");
                    if (AddColumnIfMissing("ClashZones", "WallCenterlinePointZ", "REAL DEFAULT 0.0", transaction))
                        _logger("[SQLite] ✅ Added WallCenterlinePointZ column to ClashZones (pre-calculated wall centerline point)");
                    // ✅ PIPE DIAMETER COLUMNS: Add outer diameter and nominal diameter columns for pipes
                    if (AddColumnIfMissing("ClashZones", "MepElementOuterDiameter", "REAL DEFAULT 0.0", transaction))
                        _logger("[SQLite] ✅ Added MepElementOuterDiameter column to ClashZones (pipe outer diameter in feet)");
                    if (AddColumnIfMissing("ClashZones", "MepElementNominalDiameter", "REAL DEFAULT 0.0", transaction))
                        _logger("[SQLite] ✅ Added MepElementNominalDiameter column to ClashZones (pipe nominal diameter in feet)");
                    // ✅ DAMPER TYPE/FAMILY: Add columns for damper type name and family name (stored during refresh to avoid linked file access)
                    if (AddColumnIfMissing("ClashZones", "MepElementTypeName", "TEXT", transaction))
                        _logger("[SQLite] ✅ Added MepElementTypeName column to ClashZones (damper type name: MSD, MSFD, Standard, etc.)");
                    if (AddColumnIfMissing("ClashZones", "MepElementFamilyName", "TEXT", transaction))
                        _logger("[SQLite] ✅ Added MepElementFamilyName column to ClashZones (damper family name: Motorised Smoke Damper, etc.)");
                    
                    // ✅ SIZE PARAMETER VALUE: Add string column for Size parameter value (e.g., "20 mmø", "200 mm dia symbol")
                    // This is the exact text from the Size parameter, stored for transfer to sleeve MEP_Size parameter
                    if (AddColumnIfMissing("ClashZones", "MepElementSizeParameterValue", "TEXT DEFAULT ''", transaction))
                        _logger("[SQLite] ✅ Added MepElementSizeParameterValue column to ClashZones (Size parameter as string)");
                    
                    // ✅ MEP SYSTEM INFO: Add missing system abbreviation and formatted size columns
                    if (AddColumnIfMissing("ClashZones", "MepElementSystemAbbreviation", "TEXT", transaction))
                        _logger("[SQLite] ✅ Added MepElementSystemAbbreviation column to ClashZones");
                    if (AddColumnIfMissing("ClashZones", "MepElementFormattedSize", "TEXT", transaction))
                        _logger("[SQLite] ✅ Added MepElementFormattedSize column to ClashZones");
                    
                    if (AddColumnIfMissing("ClashZones", "IsStandardDamper", "INTEGER NOT NULL DEFAULT 0", transaction))
                        _logger("[SQLite] ✅ Added IsStandardDamper column to ClashZones");

                    // ✅ MEP METADATA: Add system type, service type, and elevation from level
                    if (AddColumnIfMissing("ClashZones", "MepSystemType", "TEXT", transaction))
                        _logger("[SQLite] ✅ Added MepSystemType column to ClashZones");
                    if (AddColumnIfMissing("ClashZones", "MepServiceType", "TEXT", transaction))
                        _logger("[SQLite] ✅ Added MepServiceType column to ClashZones");
                    if (AddColumnIfMissing("ClashZones", "ElevationFromLevel", "REAL DEFAULT 0.0", transaction))
                        _logger("[SQLite] ✅ Added ElevationFromLevel column to ClashZones");
                    AddColumnIfMissing("ClashZones", "SleeveFamilyName", "TEXT", transaction);
                    AddColumnIfMissing("ClashZones", "SleevePlacementActiveX", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "SleevePlacementActiveY", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "SleevePlacementActiveZ", "REAL", transaction);
                    // Rotated bounding box columns (NULL for axis-aligned sleeves)
                    AddColumnIfMissing("ClashZones", "RotatedBoundingBoxMinX", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "RotatedBoundingBoxMinY", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "RotatedBoundingBoxMinZ", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "RotatedBoundingBoxMaxX", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "RotatedBoundingBoxMaxY", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "RotatedBoundingBoxMaxZ", "REAL", transaction);
                    // ✅ SLEEVE CORNERS: Pre-calculated 4 corner coordinates in world space (for clustering optimization)
                    // Calculated once during individual sleeve placement, stored for reuse during clustering
                    AddColumnIfMissing("ClashZones", "SleeveCorner1X", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "SleeveCorner1Y", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "SleeveCorner1Z", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "SleeveCorner2X", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "SleeveCorner2Y", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "SleeveCorner2Z", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "SleeveCorner3X", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "SleeveCorner3Y", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "SleeveCorner3Z", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "SleeveCorner4X", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "SleeveCorner4Y", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "SleeveCorner4Z", "REAL", transaction);
                    // ✅ RCS BOUNDING BOX: Wall-aligned Relative Coordinate System bounding boxes (for walls/framing only)
                    // Stored in wall-aligned coordinates: RCS X = along wall, RCS Y = through wall, RCS Z = vertical
                    // Calculated once during individual sleeve placement, stored for reuse during clustering
                    // Eliminates need for rotation logic - bounding boxes already in wall-aligned coordinates
                    AddColumnIfMissing("ClashZones", "SleeveBoundingBoxRCS_MinX", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "SleeveBoundingBoxRCS_MinY", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "SleeveBoundingBoxRCS_MinZ", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "SleeveBoundingBoxRCS_MaxX", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "SleeveBoundingBoxRCS_MaxY", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "SleeveBoundingBoxRCS_MaxZ", "REAL", transaction);
                    AddColumnIfMissing("ClashZones", "SourceDocKey", "TEXT", transaction);
                    AddColumnIfMissing("ClashZones", "HostDocKey", "TEXT", transaction);
                    AddColumnIfMissing("ClashZones", "MepElementUniqueId", "TEXT", transaction);
                    AddColumnIfMissing("ClashZones", "IsResolvedFlag", "INTEGER NOT NULL DEFAULT 0", transaction);
                    AddColumnIfMissing("ClashZones", "IsClusterResolvedFlag", "INTEGER NOT NULL DEFAULT 0", transaction);
                    // ✅ COMBINED RESOLVED: Flag for Phase 4 combined sleeves
                    if (AddColumnIfMissing("ClashZones", "IsCombinedResolved", "INTEGER NOT NULL DEFAULT 0", transaction))
                        _logger("[SQLite] ✅ Added IsCombinedResolved column to ClashZones (for combined sleeves)");
                    AddColumnIfMissing("ClashZones", "IsClusteredFlag", "INTEGER NOT NULL DEFAULT 0", transaction);
                    // ⚠️ IMPORTANT: MarkedForClusterProcess is nullable in the model (bool?).
                    // NULL = not yet evaluated for clustering (Refresh stage),
                    //  1   = should be clustered,
                    //  0   = explicitly skip clustering.
                    // Therefore the database column MUST allow NULL; using NOT NULL here causes
                    // inserts during Refresh to fail with \"NOT NULL constraint failed\".
                    AddColumnIfMissing("ClashZones", "MarkedForClusterProcess", "INTEGER", transaction);
                    AddColumnIfMissing("ClashZones", "AfterClusterSleeveId", "INTEGER NOT NULL DEFAULT -1", transaction);
                    // ✅ COMBINED RESOLVED: Instance ID for the combined sleeve if this zone is part of one
                    AddColumnIfMissing("ClashZones", "CombinedClusterSleeveInstanceId", "INTEGER", transaction);
                    // ✅ NOTE: Cluster bounding boxes are stored in ClusterSleeves table, not duplicated in ClashZones
                    // Cleanup service queries ClusterSleeves directly for cluster bounding boxes
                    AddColumnIfMissing("ClashZones", "HasDamperNearbyFlag", "INTEGER NOT NULL DEFAULT 0", transaction);
                    AddColumnIfMissing("ClashZones", "IsCurrentClashFlag", "INTEGER NOT NULL DEFAULT 0", transaction);
                    // ✅ SESSION FLAG: Track zones ready for placement in current refresh session
                    // Replaces timestamp-based filtering - more reliable for session boundaries
                    AddColumnIfMissing("ClashZones", "ReadyForPlacementFlag", "INTEGER NOT NULL DEFAULT 0", transaction);
                    // ✅ CRITICAL: Add thickness columns for depth calculation (for existing databases)
                    // Note: These are also in the initial CREATE TABLE for new databases
                    if (AddColumnIfMissing("ClashZones", "StructuralThickness", "REAL DEFAULT 0.0", transaction))
                        _logger("[SQLite] ✅ Added StructuralThickness column to ClashZones");
                    if (AddColumnIfMissing("ClashZones", "WallThickness", "REAL DEFAULT 0.0", transaction))
                        _logger("[SQLite] ✅ Added WallThickness column to ClashZones");
                    if (AddColumnIfMissing("ClashZones", "FramingThickness", "REAL DEFAULT 0.0", transaction))
                        _logger("[SQLite] ✅ Added FramingThickness column to ClashZones");
                    
                    // ✅ OOP METHOD: Add damper connector detection columns (for connector-based clearance logic)
                    if (AddColumnIfMissing("ClashZones", "HasMepConnector", "INTEGER DEFAULT 0", transaction))
                        _logger("[SQLite] ✅ Added HasMepConnector column to ClashZones (0=false, 1=true)");
                    if (AddColumnIfMissing("ClashZones", "DamperConnectorSide", "TEXT DEFAULT ''", transaction))
                        _logger("[SQLite] ✅ Added DamperConnectorSide column to ClashZones (Left/Right/Top/Bottom)");
                    
                    // ✅ OOP METHOD: Add insulation detection columns (for insulation-aware sizing)
                    if (AddColumnIfMissing("ClashZones", "IsInsulated", "INTEGER DEFAULT 0", transaction))
                        _logger("[SQLite] ✅ Added IsInsulated column to ClashZones (0=false, 1=true)");
                    if (AddColumnIfMissing("ClashZones", "InsulationThickness", "REAL DEFAULT 0.0", transaction))
                        _logger("[SQLite] ✅ Added InsulationThickness column to ClashZones (thickness in feet)");
                    
                    // ✅ PARAMETER VALUES: Store parameter values as JSON for SleeveSnapshots
                    // These are collected during refresh and needed when snapshots are saved after placement
                    if (AddColumnIfMissing("ClashZones", "MepParameterValuesJson", "TEXT", transaction))
                        _logger("[SQLite] ✅ Added MepParameterValuesJson column to ClashZones (stores MEP parameter values as JSON)");
                    if (AddColumnIfMissing("ClashZones", "HostParameterValuesJson", "TEXT", transaction))
                        _logger("[SQLite] ✅ Added HostParameterValuesJson column to ClashZones (stores Host parameter values as JSON)");
                    
                    // ✅ COMBO FLAG: Add IsFilterComboNew column to FileCombos table (defaults to 0=false for existing combos)
                    if (AddColumnIfMissing("FileCombos", "IsFilterComboNew", "INTEGER NOT NULL DEFAULT 0", transaction))
                        _logger("[SQLite] ✅ Added IsFilterComboNew column to FileCombos (existing combos default to 0=false)");
                    
                    // ✅ R-TREE: Populate R-tree index from existing ClashZones (if enabled and table exists)
                    PopulateRTreeFromExistingData(transaction);
                    
                    // ✅ PHASE 2: Add CreatedAt and UpdatedAt columns to FileCombos table for consistency
                    if (AddColumnIfMissing("FileCombos", "CreatedAt", "DATETIME NOT NULL DEFAULT (datetime('now', '+5 hours', '+30 minutes'))", transaction))
                        _logger("[SQLite] ✅ Added CreatedAt column to FileCombos");
                    if (AddColumnIfMissing("FileCombos", "UpdatedAt", "DATETIME NOT NULL DEFAULT (datetime('now', '+5 hours', '+30 minutes'))", transaction))
                        _logger("[SQLite] ✅ Added UpdatedAt column to FileCombos");
                    
                    // ✅ BACKWARD COMPATIBILITY: Keep IsFilterComboNew in Filters table for migration (deprecated, use FileCombos.IsFilterComboNew instead)
                    if (AddColumnIfMissing("Filters", "IsFilterComboNew", "INTEGER NOT NULL DEFAULT 0", transaction))
                        _logger("[SQLite] ✅ Added IsFilterComboNew column to Filters (deprecated - use FileCombos.IsFilterComboNew instead)");
                    
                    // ✅ PHASE 2: Add UI state columns to Filters table
                    // ✅ REMOVED: SelectedHostElementTypes - using SelectedHostCategories only
                    // ✅ STANDARDIZED: Use SelectedHostCategories (removed duplicate SelectedHostElementTypes)
                    if (AddColumnIfMissing("Filters", "SelectedHostCategories", "TEXT", transaction))
                        _logger("[SQLite] ✅ Added SelectedHostCategories column to Filters (JSON array)");
                    if (AddColumnIfMissing("Filters", "OpeningSettings", "TEXT", transaction))
                        _logger("[SQLite] ✅ Added OpeningSettings column to Filters (JSON) - stores clearance settings per filter");
                    // ✅ COMBO VALIDATION: Add file keys and reference category for combo validation
                    if (AddColumnIfMissing("Filters", "ReferenceDocKey", "TEXT", transaction))
                        _logger("[SQLite] ✅ Added ReferenceDocKey column to Filters (for combo validation)");
                    if (AddColumnIfMissing("Filters", "HostDocKey", "TEXT", transaction))
                        _logger("[SQLite] ✅ Added HostDocKey column to Filters (for combo validation)");
                    if (AddColumnIfMissing("Filters", "ReferenceCategory", "TEXT", transaction))
                        _logger("[SQLite] ✅ Added ReferenceCategory column to Filters (MEP category for combo validation)");
                    
                    // ✅ CONSOLIDATED: Add JSON array columns (primary storage for multiple files)
                    // These will be populated from DocKey columns as migration
                    if (AddColumnIfMissing("Filters", "SelectedMepCategoryNames", "TEXT", transaction))
                        _logger("[SQLite] ✅ Added SelectedMepCategoryNames column to Filters (JSON array of MEP categories)");
                    if (AddColumnIfMissing("Filters", "SelectedReferenceFiles", "TEXT", transaction))
                        _logger("[SQLite] ✅ Added SelectedReferenceFiles column to Filters (JSON array - migrated from ReferenceDocKey)");
                    if (AddColumnIfMissing("Filters", "SelectedHostFiles", "TEXT", transaction))
                        _logger("[SQLite] ✅ Added SelectedHostFiles column to Filters (JSON array - migrated from HostDocKey)");

                    // ✅ SCHEMA REDESIGN: Replace bloated OpeningSettings JSON with simple AdoptToDocumentFlag
                    // OpeningSettings was storing 14+ properties; we only need the AdoptToDocument flag
                    if (AddColumnIfMissing("Filters", "AdoptToDocumentFlag", "INTEGER DEFAULT 1", transaction))
                        _logger("[SQLite] ✅ Added AdoptToDocumentFlag column to Filters (boolean: 0=false, 1=true)");

                    // ✅ MIGRATION: Populate JSON array columns from existing DocKey columns
                    ExecuteCommand(@"
                        UPDATE Filters
                        SET SelectedReferenceFiles = json_array(ReferenceDocKey)
                        WHERE SelectedReferenceFiles IS NULL AND ReferenceDocKey IS NOT NULL", transaction);
                    _logger("[SQLite] ✅ Migrated ReferenceDocKey values to SelectedReferenceFiles JSON array");
                    
                    ExecuteCommand(@"
                        UPDATE Filters
                        SET SelectedHostFiles = json_array(HostDocKey)
                        WHERE SelectedHostFiles IS NULL AND HostDocKey IS NOT NULL", transaction);
                    _logger("[SQLite] ✅ Migrated HostDocKey values to SelectedHostFiles JSON array");
                    
                    ExecuteCommand(@"
                        UPDATE Filters
                        SET SelectedMepCategoryNames = json_array(ReferenceCategory)
                        WHERE SelectedMepCategoryNames IS NULL AND ReferenceCategory IS NOT NULL", transaction);
                    _logger("[SQLite] ✅ Migrated ReferenceCategory values to SelectedMepCategoryNames JSON array");

                    // ✅ DEPRECATION: OpeningSettings column is no longer used - AdoptToDocumentFlag is the replacement
                    // No migration needed - OpeningSettings was never reliably populated

                    AddColumnIfMissing("Conditions", "CombinedKey", "TEXT", transaction);
                    AddColumnIfMissing("Conditions", "PipesNormal", "REAL", transaction);
                    AddColumnIfMissing("Conditions", "PipesInsulated", "REAL", transaction);
                    AddColumnIfMissing("Conditions", "UseNominalDiameterForPipes", "INTEGER", transaction);
                    AddColumnIfMissing("Conditions", "DuctAccessoryMepNormal", "REAL", transaction);
                    AddColumnIfMissing("Conditions", "DuctAccessoryMepInsulated", "REAL", transaction);
                    AddColumnIfMissing("Conditions", "DuctAccessoryOtherNormal", "REAL", transaction);
                    AddColumnIfMissing("Conditions", "DuctAccessoryOtherInsulated", "REAL", transaction);
                    AddColumnIfMissing("Conditions", "CableTrayTopInsulated", "REAL", transaction);
                    AddColumnIfMissing("Conditions", "CableTrayOtherInsulated", "REAL", transaction);
                    AddColumnIfMissing("Conditions", "OpeningPrefs", "TEXT", transaction);
                    AddColumnIfMissing("Conditions", "HorizontalLevel", "TEXT", transaction);
                    AddColumnIfMissing("Conditions", "VerticalLevel", "TEXT", transaction);
                    AddColumnIfMissing("Conditions", "CreationMode", "TEXT", transaction);
                    AddColumnIfMissing("Conditions", "JoinOpeningsDistanceMm", "REAL", transaction);
                    AddColumnIfMissing("Conditions", "IgnoreArchitecturalFloors", "INTEGER", transaction);
                    AddColumnIfMissing("Conditions", "CircularToRectangularThresholdMm", "REAL", transaction);

                    AddColumnIfMissing("SleeveSnapshots", "SourceType", "TEXT NOT NULL DEFAULT 'Individual'", transaction);
                    AddColumnIfMissing("SleeveSnapshots", "FilterId", "INTEGER", transaction);
                    AddColumnIfMissing("SleeveSnapshots", "ComboId", "INTEGER", transaction);
                    AddColumnIfMissing("SleeveSnapshots", "MepElementIdsJson", "TEXT", transaction);
                    AddColumnIfMissing("SleeveSnapshots", "HostElementIdsJson", "TEXT", transaction);
                    AddColumnIfMissing("SleeveSnapshots", "MepParametersJson", "TEXT", transaction);
                    AddColumnIfMissing("SleeveSnapshots", "HostParametersJson", "TEXT", transaction);
                    AddColumnIfMissing("SleeveSnapshots", "SourceDocKeysJson", "TEXT", transaction);
                    AddColumnIfMissing("SleeveSnapshots", "HostDocKeysJson", "TEXT", transaction);
                    AddColumnIfMissing("SleeveSnapshots", "CreatedAt", "DATETIME NOT NULL DEFAULT (datetime('now', '+5 hours', '+30 minutes'))", transaction);
                    AddColumnIfMissing("SleeveSnapshots", "UpdatedAt", "DATETIME NOT NULL DEFAULT (datetime('now', '+5 hours', '+30 minutes'))", transaction);
                    AddColumnIfMissing("SleeveSnapshots", "ClashZoneGuid", "TEXT", transaction); // ✅ NEW: Add ClashZoneGuid column
                    
                    // ✅ CLUSTER GUID: Add deterministic GUID column to ClusterSleeves for proper upserting
                    if (AddColumnIfMissing("ClusterSleeves", "ClusterGuid", "TEXT", transaction))
                        _logger("[SQLite] ✅ Added ClusterGuid column to ClusterSleeves (deterministic GUID for upserting)");

                    // ✅ DATABASE GUID MANAGEMENT: Create indexes for GUID after column is added
                    // These indexes are created here (not in EnsureSchemaCreated) to ensure ClashZoneGuid column exists first
                    if (ColumnExists("ClashZones", "ClashZoneGuid", transaction))
                    {
                        ExecuteCommand("CREATE UNIQUE INDEX IF NOT EXISTS idx_clashzones_guid_unique ON ClashZones(ClashZoneGuid) WHERE ClashZoneGuid != ''", transaction);
                        ExecuteCommand("CREATE INDEX IF NOT EXISTS idx_clashzones_guid ON ClashZones(ClashZoneGuid)", transaction);
                    }
                    
                    // ✅ SNAPSHOT GUID INDEX: Create index for SleeveSnapshots GUID
                    if (ColumnExists("SleeveSnapshots", "ClashZoneGuid", transaction))
                    {
                        ExecuteCommand("CREATE INDEX IF NOT EXISTS idx_sleevesnapshots_guid ON SleeveSnapshots(ClashZoneGuid)", transaction);
                    }
                    
                    // ✅ CLUSTER GUID INDEX: Create index for ClusterSleeves GUID
                    if (ColumnExists("ClusterSleeves", "ClusterGuid", transaction))
                    {
                        ExecuteCommand("CREATE UNIQUE INDEX IF NOT EXISTS idx_clustersleeves_guid_unique ON ClusterSleeves(ClusterGuid) WHERE ClusterGuid != '' AND ClusterGuid IS NOT NULL", transaction);
                        ExecuteCommand("CREATE INDEX IF NOT EXISTS idx_clustersleeves_guid ON ClusterSleeves(ClusterGuid)", transaction);
                    }

                    // ✅ OPTION 4 IMPLEMENTATION: Add triggers, constraints, and views for sophisticated flag management
                    // ✅ TEMPORARILY DISABLED: Triggers causing SQL logic errors - will re-enable with compatible syntax
                    // EnsureFlagManagementTriggers(transaction);
                    EnsureFlagManagementConstraints(transaction);
                    
                    // ✅ TEMPORARILY DISABLED: Views causing SQL logic errors - will re-enable after fixing column references
                    // EnsureFlagManagementViews(transaction);
                    
                    // ✅ DATABASE GUID MANAGEMENT: Add views for GUID lookups
                    // ✅ TEMPORARILY DISABLED: Views causing SQL logic errors
                    // EnsureGuidManagementViews(transaction);

                    // ✅ COMBINED SLEEVES: EnsureCombinedSleevesTables is now called EARLIER (before AddColumnIfMissing)
                    // to fix "no such table: CombinedSleeves" error on fresh installations

                    // ✅ MIGRATION: Make ComboId and FilterId nullable for cross-filter support
                    // SQLite doesn't support ALTER COLUMN, so we need to check if FK constraints exist
                    // and recreate the table if needed
                    try
                    {
                        using (var checkCmd = _connection.CreateCommand())
                        {
                            checkCmd.Transaction = transaction;
                            
                            // ✅ CLEANUP: First, check if orphaned CombinedSleeves_Old exists from failed migration
                            checkCmd.CommandText = "SELECT name FROM sqlite_master WHERE type='table' AND name='CombinedSleeves_Old'";
                            var oldTableExists = checkCmd.ExecuteScalar() != null;
                            
                            if (oldTableExists)
                            {
                                _logger("[SQLite] 🧹 Found orphaned CombinedSleeves_Old table from failed migration - cleaning up");
                                
                                // ✅ DEFENSIVE: First check if CombinedSleeves table exists before querying it
                                checkCmd.CommandText = "SELECT name FROM sqlite_master WHERE type='table' AND name='CombinedSleeves'";
                                var combinedSleevesExists = checkCmd.ExecuteScalar() != null;
                                
                                if (combinedSleevesExists)
                                {
                                    // Check if current CombinedSleeves table exists and has data
                                    checkCmd.CommandText = "SELECT COUNT(*) FROM CombinedSleeves";
                                    var currentCount = Convert.ToInt32(checkCmd.ExecuteScalar());
                                    
                                    checkCmd.CommandText = "SELECT COUNT(*) FROM CombinedSleeves_Old";
                                    var oldCount = Convert.ToInt32(checkCmd.ExecuteScalar());
                                    
                                    // If old table has data but current doesn't, restore from old
                                    if (oldCount > 0 && currentCount == 0)
                                    {
                                        _logger($"[SQLite] 🔄 Restoring {oldCount} records from CombinedSleeves_Old");
                                        ExecuteCommand("INSERT INTO CombinedSleeves SELECT * FROM CombinedSleeves_Old", transaction);
                                    }
                                }
                                else
                                {
                                    _logger("[SQLite] ⚠️ CombinedSleeves table doesn't exist - calling EnsureCombinedSleevesTables first");
                                    EnsureCombinedSleevesTables(transaction);
                                }
                                
                                // Drop the orphaned table
                                ExecuteCommand("DROP TABLE CombinedSleeves_Old", transaction);
                                _logger("[SQLite] ✅ Cleaned up orphaned CombinedSleeves_Old table");
                            }
                            
                            // Now check if migration is needed
                            checkCmd.CommandText = "SELECT sql FROM sqlite_master WHERE type='table' AND name='CombinedSleeves'";
                            var tableSql = checkCmd.ExecuteScalar()?.ToString() ?? "";
                            
                            // Check if table has FK constraints (old schema)
                            if (tableSql.Contains("FOREIGN KEY(ComboId)") || tableSql.Contains("FOREIGN KEY(FilterId)"))
                            {
                                _logger("[SQLite] 🔄 Migrating CombinedSleeves table to remove FK constraints (cross-filter support)");
                                
                                // Step 1: Rename old table
                                ExecuteCommand("ALTER TABLE CombinedSleeves RENAME TO CombinedSleeves_Old", transaction);
                                
                                // Step 2: Create new table without FK constraints (calls EnsureCombinedSleevesTables again)
                                EnsureCombinedSleevesTables(transaction);
                                
                                // Step 3: Copy data from old table (if any columns match)
                                try
                                {
                                    ExecuteCommand(@"
                                        INSERT INTO CombinedSleeves 
                                        SELECT * FROM CombinedSleeves_Old", transaction);
                                }
                                catch
                                {
                                    // If SELECT * fails (schema mismatch), try column-by-column
                                    _logger("[SQLite] ⚠️ Schema mismatch - attempting selective column copy");
                                    ExecuteCommand(@"
                                        INSERT INTO CombinedSleeves (
                                            CombinedInstanceId, ComboId, FilterId, Categories,
                                            BoundingBoxMinX, BoundingBoxMinY, BoundingBoxMinZ,
                                            BoundingBoxMaxX, BoundingBoxMaxY, BoundingBoxMaxZ,
                                            CombinedWidth, CombinedHeight, CombinedDepth,
                                            PlacementX, PlacementY, PlacementZ
                                        )
                                        SELECT 
                                            CombinedInstanceId, ComboId, FilterId, Categories,
                                            BoundingBoxMinX, BoundingBoxMinY, BoundingBoxMinZ,
                                            BoundingBoxMaxX, BoundingBoxMaxY, BoundingBoxMaxZ,
                                            CombinedWidth, CombinedHeight, CombinedDepth,
                                            PlacementX, PlacementY, PlacementZ
                                        FROM CombinedSleeves_Old", transaction);
                                }
                                
                                // Step 4: Drop old table
                                ExecuteCommand("DROP TABLE CombinedSleeves_Old", transaction);
                                
                                _logger("[SQLite] ✅ Successfully migrated CombinedSleeves table");
                            }
                        }
                    }
                    catch (Exception migEx)
                    {
                        _logger($"[SQLite] ⚠️ CombinedSleeves migration skipped or failed: {migEx.Message}");
                        
                        // ✅ CRITICAL: If migration fails, try to clean up orphaned table to prevent future errors
                        try
                        {
                            using (var cleanupCmd = _connection.CreateCommand())
                            {
                                cleanupCmd.Transaction = transaction;
                                cleanupCmd.CommandText = "SELECT name FROM sqlite_master WHERE type='table' AND name='CombinedSleeves_Old'";
                                if (cleanupCmd.ExecuteScalar() != null)
                                {
                                    ExecuteCommand("DROP TABLE CombinedSleeves_Old", transaction);
                                    _logger("[SQLite] 🧹 Cleaned up CombinedSleeves_Old after migration failure");
                                }
                            }
                        }
                        catch (Exception cleanupEx)
                        {
                            _logger($"[SQLite] ⚠️ Failed to cleanup CombinedSleeves_Old: {cleanupEx.Message}");
                        }
                        // Non-fatal - table might not exist yet or already migrated
                    }

                    // ✅ BATCH V2: Ensure ConstituentZoneGuids column exists in ClusterSleeves_v2
                    // This column is critical for cluster placement - stores GUIDs of constituent ClashZones
                    AddColumnIfMissing("ClusterSleeves_v2", "ConstituentZoneGuids", "TEXT", transaction);
                    
                    // ✅ USER REQUEST: Add MepSystemName to ClashZones
                    AddColumnIfMissing("ClashZones", "MepSystemName", "TEXT", transaction);
                    
                    // ✅ USER REQUEST: Add WallDirectionType for wall-aligned rotation logic
                    if (AddColumnIfMissing("ClashZones", "WallDirectionType", "TEXT", transaction))
                        _logger("[SQLite] ✅ Added WallDirectionType column to ClashZones");

                    transaction.Commit();
                }
            }
            catch (Exception ex)
            {
                _logger($"[SQLite] ❌ Schema upgrade failed: {ex.Message}");
                _logger($"[SQLite] Stack trace: {ex.StackTrace}");
                throw;
            }
        }

        private bool AddColumnIfMissing(string tableName, string columnName, string columnDefinition, SQLiteTransaction transaction)
        {
            if (ColumnExists(tableName, columnName, transaction))
            {
                return false; // Column already exists
            }

            try
            {
                using (var cmd = _connection.CreateCommand())
                {
                    cmd.Transaction = transaction;
                    cmd.CommandText = $"ALTER TABLE {tableName} ADD COLUMN {columnName} {columnDefinition};";
                    cmd.ExecuteNonQuery();
                    _logger($"[SQLite] ✅ Added column '{columnName}' to table '{tableName}'");
                    return true; // Column was added
                }
            }
            catch (Exception ex)
            {
                _logger($"[SQLite] ❌ Failed to add column '{columnName}' to table '{tableName}': {ex.Message}");
                throw; // Re-throw to ensure transaction rollback
            }
        }

        private bool ColumnExists(string tableName, string columnName, SQLiteTransaction transaction)
        {
            using (var cmd = _connection.CreateCommand())
            {
                cmd.Transaction = transaction;
                cmd.CommandText = $"PRAGMA table_info({tableName});";
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        var existingName = reader["name"]?.ToString();
                        if (string.Equals(existingName, columnName, StringComparison.OrdinalIgnoreCase))
                        {
                            return true;
                        }
                    }
                }
            }

            return false;
        }

        private void EnsureSleeveSnapshotTable(SQLiteTransaction transaction)
        {
            ExecuteCommand(@"
                CREATE TABLE IF NOT EXISTS SleeveSnapshots (
                    SnapshotId          INTEGER PRIMARY KEY AUTOINCREMENT,
                    SleeveInstanceId    INTEGER,
                    ClusterInstanceId   INTEGER,
                    SourceType          TEXT NOT NULL DEFAULT 'Individual',
                    FilterId            INTEGER,
                    ComboId             INTEGER,
                    MepElementIdsJson   TEXT,
                    HostElementIdsJson  TEXT,
                    MepParametersJson   TEXT,
                    HostParametersJson  TEXT,
                    SourceDocKeysJson   TEXT,
                    HostDocKeysJson     TEXT,
                    ClashZoneGuid       TEXT,
                    CreatedAt           DATETIME NOT NULL DEFAULT (datetime('now', '+5 hours', '+30 minutes')),
                    UpdatedAt           DATETIME NOT NULL DEFAULT (datetime('now', '+5 hours', '+30 minutes')),
                    FOREIGN KEY(FilterId) REFERENCES Filters(FilterId) ON DELETE CASCADE,
                    FOREIGN KEY(ComboId) REFERENCES FileCombos(ComboId) ON DELETE CASCADE
                )", transaction);

            ExecuteCommand("CREATE INDEX IF NOT EXISTS idx_sleevesnapshots_sleeve ON SleeveSnapshots(SleeveInstanceId)", transaction);
            ExecuteCommand("CREATE INDEX IF NOT EXISTS idx_sleevesnapshots_cluster ON SleeveSnapshots(ClusterInstanceId)", transaction);
            ExecuteCommand("CREATE INDEX IF NOT EXISTS idx_sleevesnapshots_guid ON SleeveSnapshots(ClashZoneGuid)", transaction);

            // ✅ MIGRATION: Add CombinedInstanceId column if missing (Phase 4 Combined Sleeves)
            AddColumnIfMissing("SleeveSnapshots", "CombinedInstanceId", "INTEGER", transaction);
            ExecuteCommand("CREATE INDEX IF NOT EXISTS idx_sleevesnapshots_combined ON SleeveSnapshots(CombinedInstanceId)", transaction);
        }

        /// <summary>
        /// ✅ PARAMETER TRANSFER OPTIMIZATION: Create ParameterTransferFlags table
        /// Stores flags for which parameters have been transferred to which sleeves
        /// Enables ultra-fast skip lookup (just check if record exists) instead of comparing values
        /// </summary>
        private void EnsureParameterTransferFlagsTable(SQLiteTransaction transaction)
        {
            // ✅ FIX: Remove FK to ClashZones to avoid mismatch errors on existing DBs
            // We only need a fast lookup table; integrity is enforced at application level.
            ExecuteCommand("DROP TABLE IF EXISTS ParameterTransferFlags", transaction);

            ExecuteCommand(@"CREATE TABLE IF NOT EXISTS ParameterTransferFlags (
                    FlagId              INTEGER PRIMARY KEY AUTOINCREMENT,
                    SleeveInstanceId    INTEGER NOT NULL,
                    ParameterName       TEXT NOT NULL,
                    TransferredAt       DATETIME NOT NULL DEFAULT (datetime('now', '+5 hours', '+30 minutes')),
                    UNIQUE(SleeveInstanceId, ParameterName)
                )", transaction);

            ExecuteCommand("CREATE INDEX IF NOT EXISTS idx_transfer_flags_sleeve ON ParameterTransferFlags(SleeveInstanceId)", transaction);
            ExecuteCommand("CREATE INDEX IF NOT EXISTS idx_transfer_flags_param ON ParameterTransferFlags(ParameterName)", transaction);
        }

        /// <summary>
        /// ✅ CATEGORY PROCESSING METADATA: Create ProcessingMarkers table
        /// Tracks the last processed sleeve count per category to enable incremental processing
        /// Instead of recounting from 1, we only append new sleeves since last mark
        /// </summary>
        private void EnsureCategoryProcessingMarkersTable(SQLiteTransaction transaction)
        {
            ExecuteCommand(@"CREATE TABLE IF NOT EXISTS CategoryProcessingMarkers (
                    MarkerId            INTEGER PRIMARY KEY AUTOINCREMENT,
                    Category            TEXT NOT NULL UNIQUE,
                    LastProcessedCount  INTEGER NOT NULL DEFAULT 0,
                    LastProcessedSleeveIds TEXT,
                    MarkedAt            DATETIME NOT NULL DEFAULT (datetime('now', '+5 hours', '+30 minutes')),
                    UpdatedAt           DATETIME NOT NULL DEFAULT (datetime('now', '+5 hours', '+30 minutes'))
                )", transaction);

            ExecuteCommand("CREATE INDEX IF NOT EXISTS idx_markers_category ON CategoryProcessingMarkers(Category)", transaction);
        }

        /// <summary>
        /// ✅ CLUSTER SLEEVE STORAGE: Create ClusterSleeves table to store cluster calculation results
        /// This enables PATH 1 (Replay) to use pre-calculated cluster data without recalculating
        /// </summary>
        private void EnsureClusterSleevesTable(SQLiteTransaction transaction)
        {
            ExecuteCommand(@"CREATE TABLE IF NOT EXISTS ClusterSleeves (
                    ClusterSleeveId      INTEGER PRIMARY KEY AUTOINCREMENT,
                    ClusterInstanceId   INTEGER NOT NULL UNIQUE,
                    ComboId             INTEGER NOT NULL,
                    FilterId            INTEGER NOT NULL,
                    Category            TEXT NOT NULL,
                    BoundingBoxMinX     REAL NOT NULL,
                    BoundingBoxMinY     REAL NOT NULL,
                    BoundingBoxMinZ     REAL NOT NULL,
                    BoundingBoxMaxX     REAL NOT NULL,
                    BoundingBoxMaxY     REAL NOT NULL,
                    BoundingBoxMaxZ     REAL NOT NULL,
                    ClusterWidth        REAL NOT NULL,
                    ClusterHeight       REAL NOT NULL,
                    ClusterDepth        REAL NOT NULL,
                    RotationAngleDeg    REAL DEFAULT 0.0,
                    IsRotated           INTEGER NOT NULL DEFAULT 0,
                    PlacementX          REAL NOT NULL,
                    PlacementY          REAL NOT NULL,
                    PlacementZ          REAL NOT NULL,
                    HostType            TEXT,
                    HostOrientation     TEXT,
                    ClashZoneIdsJson    TEXT NOT NULL,
                    ClashZoneGuids      TEXT,
                    MepSizes            TEXT,
                    MepSystemNames      TEXT,
                    MepElementIds       TEXT,
                    Corner1X            REAL DEFAULT 0.0,
                    Corner1Y            REAL DEFAULT 0.0,
                    Corner1Z            REAL DEFAULT 0.0,
                    Corner2X            REAL DEFAULT 0.0,
                    Corner2Y            REAL DEFAULT 0.0,
                    Corner2Z            REAL DEFAULT 0.0,
                    Corner3X            REAL DEFAULT 0.0,
                    Corner3Y            REAL DEFAULT 0.0,
                    Corner3Z            REAL DEFAULT 0.0,
                    Corner4X            REAL DEFAULT 0.0,
                    Corner4Y            REAL DEFAULT 0.0,
                    Corner4Z            REAL DEFAULT 0.0,
                    CreatedAt           DATETIME NOT NULL DEFAULT (datetime('now', '+5 hours', '+30 minutes')),
                    UpdatedAt           DATETIME NOT NULL DEFAULT (datetime('now', '+5 hours', '+30 minutes')),
                    FOREIGN KEY(ComboId) REFERENCES FileCombos(ComboId) ON DELETE CASCADE,
                    FOREIGN KEY(FilterId) REFERENCES Filters(FilterId) ON DELETE CASCADE
                )", transaction);

            // Create indexes for fast lookups
            ExecuteCommand("CREATE INDEX IF NOT EXISTS idx_clustersleeves_instance ON ClusterSleeves(ClusterInstanceId)", transaction);
            ExecuteCommand("CREATE INDEX IF NOT EXISTS idx_clustersleeves_combo ON ClusterSleeves(ComboId)", transaction);
            ExecuteCommand("CREATE INDEX IF NOT EXISTS idx_clustersleeves_filter ON ClusterSleeves(FilterId)", transaction);
            ExecuteCommand("CREATE INDEX IF NOT EXISTS idx_clustersleeves_category ON ClusterSleeves(Category)", transaction);
            ExecuteCommand("CREATE INDEX IF NOT EXISTS idx_clustersleeves_combo_category ON ClusterSleeves(ComboId, Category)", transaction);
            
            // ✅ MIGRATION: Add comma-separated value columns if they don't exist
            AddColumnIfMissing("ClusterSleeves", "ClashZoneGuids", "TEXT", transaction);
            AddColumnIfMissing("ClusterSleeves", "MepSizes", "TEXT", transaction);
            AddColumnIfMissing("ClusterSleeves", "MepSystemNames", "TEXT", transaction);
            AddColumnIfMissing("ClusterSleeves", "MepServiceTypes", "TEXT", transaction); // ✅ ADDED: MepServiceTypes
            AddColumnIfMissing("ClusterSleeves", "MepElementIds", "TEXT", transaction);
            AddColumnIfMissing("ClusterSleeves", "SleeveFamilyName", "TEXT", transaction);
            
            // ✅ MIGRATION: Add Corner columns if they don't exist (Phase 3 Persistence)
            AddColumnIfMissing("ClusterSleeves", "Corner1X", "REAL DEFAULT 0.0", transaction);
            AddColumnIfMissing("ClusterSleeves", "Corner1Y", "REAL DEFAULT 0.0", transaction);
            AddColumnIfMissing("ClusterSleeves", "Corner1Z", "REAL DEFAULT 0.0", transaction);
            AddColumnIfMissing("ClusterSleeves", "Corner2X", "REAL DEFAULT 0.0", transaction);
            AddColumnIfMissing("ClusterSleeves", "Corner2Y", "REAL DEFAULT 0.0", transaction);
            AddColumnIfMissing("ClusterSleeves", "Corner2Z", "REAL DEFAULT 0.0", transaction);
            AddColumnIfMissing("ClusterSleeves", "Corner3X", "REAL DEFAULT 0.0", transaction);
            AddColumnIfMissing("ClusterSleeves", "Corner3Y", "REAL DEFAULT 0.0", transaction);
            AddColumnIfMissing("ClusterSleeves", "Corner3Z", "REAL DEFAULT 0.0", transaction);
            AddColumnIfMissing("ClusterSleeves", "Corner4X", "REAL DEFAULT 0.0", transaction);
            AddColumnIfMissing("ClusterSleeves", "Corner4Y", "REAL DEFAULT 0.0", transaction);
            AddColumnIfMissing("ClusterSleeves", "Corner4Z", "REAL DEFAULT 0.0", transaction);
            
            // ✅ COMBINED SLEEVE TRACKING: Instance ID for the combined sleeve if this cluster is part of one
            // NOTE: IsCombinedResolved is NOT needed here - only in ClashZones
            // But CombinedClusterSleeveInstanceId IS needed for parameter transfer lookup
            AddColumnIfMissing("ClusterSleeves", "CombinedClusterSleeveInstanceId", "INTEGER DEFAULT -1", transaction);

            // ✅ PERSISTENCE FIX: Add BoundingBox columns to Legacy Table
            // Required for Stage 2 Cleanup which queries these columns
            AddColumnIfMissing("ClusterSleeves", "BoundingBoxMinX", "REAL DEFAULT 0.0", transaction);
            AddColumnIfMissing("ClusterSleeves", "BoundingBoxMinY", "REAL DEFAULT 0.0", transaction);
            AddColumnIfMissing("ClusterSleeves", "BoundingBoxMinZ", "REAL DEFAULT 0.0", transaction);
            AddColumnIfMissing("ClusterSleeves", "BoundingBoxMaxX", "REAL DEFAULT 0.0", transaction);
            AddColumnIfMissing("ClusterSleeves", "BoundingBoxMaxY", "REAL DEFAULT 0.0", transaction);
            AddColumnIfMissing("ClusterSleeves", "BoundingBoxMaxZ", "REAL DEFAULT 0.0", transaction);
        }

        /// <summary>
        /// ✅ OPTION 4 IMPLEMENTATION: Ensure database triggers for automatic flag management
        /// Triggers automatically sync flags with sleeve IDs and compute SleeveState
        /// </summary>
        private void EnsureFlagManagementTriggers(SQLiteTransaction transaction)
        {
            try
            {
                // ✅ FIXED: Simplified trigger syntax for SQLite compatibility
                // Trigger 1: Auto-update SleeveState when flags change
                // Note: Using simpler WHEN clause to avoid SQL logic errors
                
                // ✅ CRITICAL: Drop existing trigger first to ensure update
                ExecuteCommand("DROP TRIGGER IF EXISTS update_sleeve_state_on_flags", transaction);
                
                ExecuteCommand(@"
                    CREATE TRIGGER IF NOT EXISTS update_sleeve_state_on_flags
                    AFTER UPDATE OF IsResolvedFlag, IsClusterResolvedFlag ON ClashZones
                    FOR EACH ROW
                    WHEN (OLD.IsResolvedFlag IS DISTINCT FROM NEW.IsResolvedFlag OR 
                          OLD.IsClusterResolvedFlag IS DISTINCT FROM NEW.IsClusterResolvedFlag)
                    BEGIN
                        UPDATE ClashZones
                        SET SleeveState = CASE 
                            WHEN NEW.IsClusterResolvedFlag = 1 THEN 2
                            WHEN NEW.IsResolvedFlag = 1 THEN 1
                            ELSE 0
                        END
                        WHERE ClashZoneId = NEW.ClashZoneId
                          AND (
                              (NEW.IsClusterResolvedFlag = 1 AND SleeveState != 2) OR
                              (NEW.IsClusterResolvedFlag = 0 AND NEW.IsResolvedFlag = 1 AND SleeveState != 1) OR
                              (NEW.IsClusterResolvedFlag = 0 AND NEW.IsResolvedFlag = 0 AND SleeveState != 0)
                          );
                    END;", transaction);
            }
            catch (Exception ex)
            {
                // ✅ FALLBACK: If trigger creation fails, log and continue (triggers are optional optimization)
                _logger($"[SQLite] ⚠️ Could not create update_sleeve_state_on_flags trigger: {ex.Message}");
            }

            try
            {
                // Trigger 2: Auto-update flags and SleeveState when sleeve IDs change
                
                // ✅ CRITICAL: Drop existing trigger first to ensure update
                ExecuteCommand("DROP TRIGGER IF EXISTS sync_flags_from_ids", transaction);
                
                ExecuteCommand(@"
                    CREATE TRIGGER IF NOT EXISTS sync_flags_from_ids
                    AFTER UPDATE OF SleeveInstanceId, ClusterInstanceId ON ClashZones
                    FOR EACH ROW
                    WHEN (OLD.SleeveInstanceId IS DISTINCT FROM NEW.SleeveInstanceId OR 
                          OLD.ClusterInstanceId IS DISTINCT FROM NEW.ClusterInstanceId)
                    BEGIN
                        UPDATE ClashZones
                        SET 
                            IsResolvedFlag = CASE WHEN NEW.SleeveInstanceId > 0 THEN 1 ELSE 0 END,
                            IsClusterResolvedFlag = CASE WHEN NEW.ClusterInstanceId > 0 THEN 1 ELSE 0 END,
                            SleeveState = CASE 
                                WHEN NEW.ClusterInstanceId > 0 THEN 2
                                WHEN NEW.SleeveInstanceId > 0 THEN 1
                                ELSE 0
                            END
                        WHERE ClashZoneId = NEW.ClashZoneId
                          AND (
                              (NEW.SleeveInstanceId > 0 AND IsResolvedFlag != 1) OR
                              (NEW.SleeveInstanceId <= 0 AND IsResolvedFlag != 0) OR
                              (NEW.ClusterInstanceId > 0 AND IsClusterResolvedFlag != 1) OR
                              (NEW.ClusterInstanceId <= 0 AND IsClusterResolvedFlag != 0)
                          );
                    END;", transaction);
            }
            catch (Exception ex)
            {
                // ✅ FALLBACK: If trigger creation fails, log and continue (triggers are optional optimization)
                _logger($"[SQLite] ⚠️ Could not create sync_flags_from_ids trigger: {ex.Message}");
            }

            _logger("[SQLite] ✅ Flag management triggers creation attempted (non-blocking)");
        }

        /// <summary>
        /// ✅ OPTION 4 IMPLEMENTATION: Ensure CHECK constraints for flag consistency
        /// Constraints prevent invalid flag states at the database level
        /// </summary>
        private void EnsureFlagManagementConstraints(SQLiteTransaction transaction)
        {
            // Note: SQLite doesn't support adding CHECK constraints via ALTER TABLE
            // We'll validate in application code, but document the constraint logic
            // For SQLite, we rely on triggers and application-level validation
            
            // Create a helper table to track constraint violations (optional)
            ExecuteCommand(@"
                CREATE TABLE IF NOT EXISTS FlagConstraintViolations (
                    ViolationId INTEGER PRIMARY KEY AUTOINCREMENT,
                    ClashZoneId INTEGER NOT NULL,
                    ViolationType TEXT NOT NULL,
                    Details TEXT,
                    DetectedAt DATETIME NOT NULL DEFAULT (datetime('now', '+5 hours', '+30 minutes')),
                    FOREIGN KEY(ClashZoneId) REFERENCES ClashZones(ClashZoneId) ON DELETE CASCADE
                )", transaction);

            _logger("[SQLite] ✅ Flag management constraint tracking table created/verified");
        }

        /// <summary>
        /// ✅ OPTION 4 IMPLEMENTATION: Ensure database views for complex flag queries
        /// Views provide optimized access to flag-related data
        /// </summary>
        private void EnsureFlagManagementViews(SQLiteTransaction transaction)
        {
            // View 1: Resolved clash zones (for quick filtering)
            ExecuteCommand(@"
                CREATE VIEW IF NOT EXISTS ResolvedClashZones AS
                SELECT cz.*
                FROM ClashZones cz
                WHERE cz.IsResolvedFlag = 1 OR cz.IsClusterResolvedFlag = 1;", transaction);

            // View 2: Unresolved clash zones (for placement operations)
            ExecuteCommand(@"
                CREATE VIEW IF NOT EXISTS UnresolvedClashZones AS
                SELECT cz.*
                FROM ClashZones cz
                WHERE cz.IsResolvedFlag = 0 AND cz.IsClusterResolvedFlag = 0;", transaction);

            // View 3: Clash zone flags summary (for reporting)
            ExecuteCommand(@"
                CREATE VIEW IF NOT EXISTS ClashZoneFlagsSummary AS
                SELECT 
                    cz.ClashZoneId,
                    cz.ClashZoneGuid,
                    cz.SleeveInstanceId,
                    cz.ClusterInstanceId,
                    cz.IsResolvedFlag,
                    cz.IsClusterResolvedFlag,
                    cz.SleeveState,
                    CASE 
                        WHEN cz.ClusterInstanceId > 0 THEN 'Cluster Resolved'
                        WHEN cz.SleeveInstanceId > 0 THEN 'Individual Resolved'
                        ELSE 'Unresolved'
                    END AS ResolutionStatus,
                    fc.LinkedFileKey,
                    fc.HostFileKey,
                    f.FilterName,
                    f.Category
                FROM ClashZones cz
                INNER JOIN FileCombos fc ON cz.ComboId = fc.ComboId
                INNER JOIN Filters f ON fc.FilterId = f.FilterId;", transaction);

            _logger("[SQLite] ✅ Flag management views created/verified");
        }

        /// <summary>
        /// ✅ DATABASE GUID MANAGEMENT: Ensure database views for GUID lookups
        /// Views provide optimized access to GUID-related data
        /// </summary>
        private void EnsureGuidManagementViews(SQLiteTransaction transaction)
        {
            // View 1: GUID lookup by MEP+Host+Point (for deterministic GUID matching)
            ExecuteCommand(@"
                CREATE VIEW IF NOT EXISTS ClashZoneGuidLookup AS
                SELECT 
                    ClashZoneGuid,
                    ClashZoneId,
                    MepElementId,
                    HostElementId,
                    IntersectionX,
                    IntersectionY,
                    IntersectionZ,
                    ComboId
                FROM ClashZones
                WHERE ClashZoneGuid != '' AND ClashZoneGuid IS NOT NULL;", transaction);

            _logger("[SQLite] ✅ GUID management views created/verified");
        }

        /// <summary>
        /// Get database connection (for direct SQL operations)
        /// </summary>
        public SQLiteConnection Connection => _connection;

        /// <summary>
        /// Database file path
        /// </summary>
        public string DatabasePath => _databasePath;

        public void Dispose()
        {
            if (!_disposed)
            {
                if (_connection != null)
                {
                    _connection.Close();
#if NET8_0_OR_GREATER
                    // ✅ CRITICAL: Clear pool for .NET 8/Microsoft.Data.Sqlite to prevent file locks
                    SQLiteConnection.ClearPool(_connection);
#endif
                    _connection.Dispose();
                }
                _disposed = true;
            }
        }
        
        /// <summary>
        /// ✅ R-TREE: Create R-tree virtual table for spatial indexing
        /// Only creates if UseRTreeDatabaseIndex flag is enabled
        /// Falls back gracefully if R-tree is not available
        /// </summary>
        private void EnsureRTreeTable(SQLiteTransaction transaction)
        {
            if (!Services.OptimizationFlags.UseRTreeDatabaseIndex)
            {
                _logger("[SQLite] R-tree database index disabled (UseRTreeDatabaseIndex=false) - skipping R-tree table creation");
                return;
            }
            
            try
            {
                // Check if R-tree is available (from CheckRTreeSupport)
                // If not available, skip creation and log warning
                bool rtreeAvailable = false;
                try
                {
                    using (var testCmd = _connection.CreateCommand())
                    {
                        testCmd.CommandText = "CREATE VIRTUAL TABLE IF NOT EXISTS _rtree_test_availability USING rtree(id INTEGER PRIMARY KEY, minX REAL, maxX REAL, minY REAL, maxY REAL, minZ REAL, maxZ REAL)";
                        testCmd.Transaction = transaction;
                        testCmd.ExecuteNonQuery();
                        
                        testCmd.CommandText = "DROP TABLE IF EXISTS _rtree_test_availability";
                        testCmd.ExecuteNonQuery();
                        
                        rtreeAvailable = true;
                    }
                }
                catch
                {
                    rtreeAvailable = false;
                }
                
                if (!rtreeAvailable)
                {
                    _logger("[SQLite] ⚠️ R-tree not available - skipping R-tree table creation (will use B-tree fallback)");
                    return;
                }
                
                // Create R-tree virtual table for spatial indexing
                // Links to ClashZones table via ClashZoneId
                ExecuteCommand(@"
                    CREATE VIRTUAL TABLE IF NOT EXISTS ClashZonesRTree USING rtree(
                        id INTEGER PRIMARY KEY,           -- ClashZoneId (links to ClashZones.ClashZoneId)
                        minX REAL, maxX REAL,            -- Bounding box X coordinates (world space)
                        minY REAL, maxY REAL,            -- Bounding box Y coordinates (world space)
                        minZ REAL, maxZ REAL             -- Bounding box Z coordinates (world space)
                    )", transaction);
                
                _logger("[SQLite] ✅ R-tree virtual table created/verified: ClashZonesRTree");
            }
            catch (Exception ex)
            {
                _logger($"[SQLite] ⚠️ Failed to create R-tree table (will use B-tree fallback): {ex.Message}");
                // Don't throw - allow fallback to B-tree indexes
            }
        }
        
        /// <summary>
        /// ✅ R-TREE MIGRATION: Populate R-tree index from existing ClashZones data
        /// Called during schema upgrade to backfill R-tree for existing databases
        /// </summary>
        private void PopulateRTreeFromExistingData(SQLiteTransaction transaction)
        {
            if (!Services.OptimizationFlags.UseRTreeDatabaseIndex)
                return;
            
            try
            {
                // Check if R-tree table exists
                using (var checkCmd = _connection.CreateCommand())
                {
                    checkCmd.Transaction = transaction;
                    checkCmd.CommandText = "SELECT name FROM sqlite_master WHERE type='table' AND name='ClashZonesRTree'";
                    var tableExists = checkCmd.ExecuteScalar() != null;
                    
                    if (!tableExists)
                    {
                        _logger("[SQLite] R-tree table does not exist - skipping population");
                        return;
                    }
                }
                
                // ✅ FIX: Check if R-tree count matches expected count (zones with valid bounding boxes)
                // If counts don't match, clear and re-populate to ensure consistency
                int rtreeCount = 0;
                int expectedCount = 0;
                
                using (var countCmd = _connection.CreateCommand())
                {
                    countCmd.Transaction = transaction;
                    
                    // Get current R-tree count
                    countCmd.CommandText = "SELECT COUNT(*) FROM ClashZonesRTree";
                    rtreeCount = Convert.ToInt32(countCmd.ExecuteScalar());
                    
                    // Get expected count (zones with valid bounding boxes)
                    // Use same validation logic as UpdateRTreeIndex: min < max and not all zeros
                    countCmd.CommandText = @"
                        SELECT COUNT(*) FROM ClashZones
                        WHERE BoundingBoxMinX IS NOT NULL 
                          AND BoundingBoxMaxX IS NOT NULL
                          AND BoundingBoxMinY IS NOT NULL 
                          AND BoundingBoxMaxY IS NOT NULL
                          AND BoundingBoxMinZ IS NOT NULL 
                          AND BoundingBoxMaxZ IS NOT NULL
                          AND BoundingBoxMinX < BoundingBoxMaxX
                          AND BoundingBoxMinY < BoundingBoxMaxY
                          AND BoundingBoxMinZ < BoundingBoxMaxZ
                          AND (BoundingBoxMinX != 0.0 OR BoundingBoxMaxX != 0.0 OR
                               BoundingBoxMinY != 0.0 OR BoundingBoxMaxY != 0.0 OR
                               BoundingBoxMinZ != 0.0 OR BoundingBoxMaxZ != 0.0)";
                    expectedCount = Convert.ToInt32(countCmd.ExecuteScalar());
                    
                    _logger($"[SQLite] R-tree check: Current={rtreeCount}, Expected={expectedCount}");
                    
                    // If counts match and R-tree has entries, skip re-population
                    if (rtreeCount == expectedCount && rtreeCount > 0)
                    {
                        _logger($"[SQLite] ✅ R-tree already populated correctly with {rtreeCount} entries - skipping");
                        return;
                    }
                    
                    // ✅ FIX: If counts don't match, clear and re-populate
                    if (rtreeCount > 0)
                    {
                        _logger($"[SQLite] ⚠️ R-tree count mismatch ({rtreeCount} vs {expectedCount}) - clearing and re-populating");
                        countCmd.CommandText = "DELETE FROM ClashZonesRTree";
                        countCmd.ExecuteNonQuery();
                    }
                }
                
                // Populate R-tree from existing ClashZones
                // Use BoundingBoxMinX/MaxX columns (stored in database, mapped from SleeveBoundingBoxMinX/MaxX)
                // Only include valid bounding boxes (min < max for each dimension and not all zeros)
                using (var populateCmd = _connection.CreateCommand())
                {
                    populateCmd.Transaction = transaction;
                    populateCmd.CommandText = @"
                        INSERT INTO ClashZonesRTree (id, minX, maxX, minY, maxY, minZ, maxZ)
                        SELECT 
                            ClashZoneId,
                            BoundingBoxMinX, BoundingBoxMaxX,
                            BoundingBoxMinY, BoundingBoxMaxY,
                            BoundingBoxMinZ, BoundingBoxMaxZ
                        FROM ClashZones
                        WHERE BoundingBoxMinX IS NOT NULL 
                          AND BoundingBoxMaxX IS NOT NULL
                          AND BoundingBoxMinY IS NOT NULL 
                          AND BoundingBoxMaxY IS NOT NULL
                          AND BoundingBoxMinZ IS NOT NULL 
                          AND BoundingBoxMaxZ IS NOT NULL
                          AND BoundingBoxMinX < BoundingBoxMaxX
                          AND BoundingBoxMinY < BoundingBoxMaxY
                          AND BoundingBoxMinZ < BoundingBoxMaxZ
                          AND (BoundingBoxMinX != 0.0 OR BoundingBoxMaxX != 0.0 OR
                               BoundingBoxMinY != 0.0 OR BoundingBoxMaxY != 0.0 OR
                               BoundingBoxMinZ != 0.0 OR BoundingBoxMaxZ != 0.0)";
                    
                    var rowsInserted = populateCmd.ExecuteNonQuery();
                    _logger($"[SQLite] ✅ Populated R-tree index with {rowsInserted} entries from existing ClashZones (expected {expectedCount})");
                }
            }
            catch (Exception ex)
            {
                _logger($"[SQLite] ⚠️ Failed to populate R-tree from existing data: {ex.Message}");
                // Don't throw - allow fallback to B-tree indexes
            }
        }
        /// <summary>
        /// Ensures Combined Sleeves tables exist (Phase 4)
        /// </summary>
        /// <summary>
        /// ✅ BATCH V2: Ensure ClusterSleeves_v2 table exists
        /// Designed for decoupled "Calculation First" workflow using GUIDs
        /// </summary>
        private void EnsureClusterSleevesV2Table(SQLiteTransaction transaction)
        {
            ExecuteCommand(@"
                CREATE TABLE IF NOT EXISTS ClusterSleeves_v2 (
                    -- Identity
                    ClusterGUID             TEXT PRIMARY KEY,
                    ClusterBatchId          TEXT NOT NULL,
                    
                    -- Placement Data (calculated in Phase 1)
                    PlacementX              REAL NOT NULL,
                    PlacementY              REAL NOT NULL,
                    PlacementZ              REAL NOT NULL,
                    ClusterWidth            REAL NOT NULL,
                    ClusterHeight           REAL NOT NULL,
                    ClusterDepth            REAL NOT NULL,
                    RotationAngleRad        REAL DEFAULT 0.0,
                    
                    -- Host/Context
                    HostElementId           INTEGER,
                    HostType                TEXT,
                    HostOrientation         TEXT,
                    Category                TEXT,
                    FamilyName              TEXT,
                    
                    -- Relationships
                    ConstituentZoneGuids    TEXT,  -- JSON array
                    ComboId                 INTEGER,
                    FilterId                INTEGER,
                    
                    -- Revit State (updated in Phase 2)
                    ClusterInstanceId       INTEGER DEFAULT -1,
                    Status                  TEXT DEFAULT 'Pending',  -- 'Pending', 'Placed', 'Failed', 'Skipped'
                    
                    -- Validation
                    ValidationStatus        TEXT DEFAULT 'Valid',
                    ValidationMessage       TEXT,
                    
                    -- Timestamps
                    CalculatedAt            DATETIME DEFAULT CURRENT_TIMESTAMP,
                    PlacedAt                DATETIME
                )", transaction);

            // Create indexes for performance
            ExecuteCommand("CREATE INDEX IF NOT EXISTS idx_clustersleevesv2_batchid ON ClusterSleeves_v2(ClusterBatchId)", transaction);
            ExecuteCommand("CREATE INDEX IF NOT EXISTS idx_clustersleevesv2_status ON ClusterSleeves_v2(Status)", transaction);
            
            // ✅ MIGRATION: Add columns to match V1 functionality
            AddColumnIfMissing("ClusterSleeves_v2", "ClashZoneIdsJson", "TEXT", transaction);
            AddColumnIfMissing("ClusterSleeves_v2", "ClashZoneGuids", "TEXT", transaction);
            AddColumnIfMissing("ClusterSleeves_v2", "MepSizes", "TEXT", transaction);
            AddColumnIfMissing("ClusterSleeves_v2", "MepSystemNames", "TEXT", transaction);
            AddColumnIfMissing("ClusterSleeves_v2", "MepServiceTypes", "TEXT", transaction);
            AddColumnIfMissing("ClusterSleeves_v2", "MepElementIds", "TEXT", transaction);
            AddColumnIfMissing("ClusterSleeves_v2", "IsRotated", "INTEGER DEFAULT 0", transaction);
            
            // ✅ CRITICAL: Ensure ClusterInstanceId exists for mapping after placement
            // This connects the Revit Element to the pre-calculated cluster data.
            AddColumnIfMissing("ClusterSleeves_v2", "ClusterInstanceId", "INTEGER DEFAULT -1", transaction);
            AddColumnIfMissing("ClusterSleeves_v2", "ComboId", "INTEGER", transaction);
            AddColumnIfMissing("ClusterSleeves_v2", "FilterId", "INTEGER", transaction);

            // ✅ BOUNDING BOXES: Add bounding box columns for Stage 2 Cleanup
            AddColumnIfMissing("ClusterSleeves_v2", "BoundingBoxMinX", "REAL DEFAULT 0.0", transaction);
            AddColumnIfMissing("ClusterSleeves_v2", "BoundingBoxMinY", "REAL DEFAULT 0.0", transaction);
            AddColumnIfMissing("ClusterSleeves_v2", "BoundingBoxMinZ", "REAL DEFAULT 0.0", transaction);
            AddColumnIfMissing("ClusterSleeves_v2", "BoundingBoxMaxX", "REAL DEFAULT 0.0", transaction);
            AddColumnIfMissing("ClusterSleeves_v2", "BoundingBoxMaxY", "REAL DEFAULT 0.0", transaction);
            AddColumnIfMissing("ClusterSleeves_v2", "BoundingBoxMaxZ", "REAL DEFAULT 0.0", transaction);
            
            // ✅ CORNERS: Add all corner columns
            AddColumnIfMissing("ClusterSleeves_v2", "Corner1X", "REAL DEFAULT 0.0", transaction);
            AddColumnIfMissing("ClusterSleeves_v2", "Corner1Y", "REAL DEFAULT 0.0", transaction);
            AddColumnIfMissing("ClusterSleeves_v2", "Corner1Z", "REAL DEFAULT 0.0", transaction);
            AddColumnIfMissing("ClusterSleeves_v2", "Corner2X", "REAL DEFAULT 0.0", transaction);
            AddColumnIfMissing("ClusterSleeves_v2", "Corner2Y", "REAL DEFAULT 0.0", transaction);
            AddColumnIfMissing("ClusterSleeves_v2", "Corner2Z", "REAL DEFAULT 0.0", transaction);
            AddColumnIfMissing("ClusterSleeves_v2", "Corner3X", "REAL DEFAULT 0.0", transaction);
            AddColumnIfMissing("ClusterSleeves_v2", "Corner3Y", "REAL DEFAULT 0.0", transaction);
            AddColumnIfMissing("ClusterSleeves_v2", "Corner3Z", "REAL DEFAULT 0.0", transaction);
            AddColumnIfMissing("ClusterSleeves_v2", "Corner4X", "REAL DEFAULT 0.0", transaction);
            AddColumnIfMissing("ClusterSleeves_v2", "Corner4Y", "REAL DEFAULT 0.0", transaction);
            AddColumnIfMissing("ClusterSleeves_v2", "Corner4Z", "REAL DEFAULT 0.0", transaction);

            // ✅ SLEEVE FAMILY NAME: Add family name column
            AddColumnIfMissing("ClusterSleeves_v2", "SleeveFamilyName", "TEXT", transaction);
        }

        private void EnsureCombinedSleevesTables(SQLiteTransaction transaction)
        {
            // ✅ CRITICAL FIX: Ensure table exists BEFORE adding columns
            // Table 1: Combined Sleeves
            // ✅ CROSS-FILTER SUPPORT: ComboId and FilterId are nullable because combined sleeves
            // can span multiple filters/combos (e.g., Pipes from Filter A + Duct Accessories from Filter B)
            ExecuteCommand(@"CREATE TABLE IF NOT EXISTS CombinedSleeves (
                    CombinedSleeveId    INTEGER PRIMARY KEY AUTOINCREMENT,
                    CombinedInstanceId  INTEGER NOT NULL UNIQUE,
                    DeterministicGuid   TEXT UNIQUE,
                    ComboId             INTEGER,
                    FilterId            INTEGER,
                    Categories          TEXT NOT NULL,
                    BoundingBoxMinX     REAL NOT NULL,
                    BoundingBoxMinY     REAL NOT NULL,
                    BoundingBoxMinZ     REAL NOT NULL,
                    BoundingBoxMaxX     REAL NOT NULL,
                    BoundingBoxMaxY     REAL NOT NULL,
                    BoundingBoxMaxZ     REAL NOT NULL,
                    CombinedWidth       REAL NOT NULL,
                    CombinedHeight      REAL NOT NULL,
                    CombinedDepth       REAL NOT NULL,
                    PlacementX          REAL NOT NULL,
                    PlacementY          REAL NOT NULL,
                    PlacementZ          REAL NOT NULL,
                    RotationAngleDeg    REAL DEFAULT 0.0,
                    HostType            TEXT,
                    HostOrientation     TEXT,
                    Corner1X            REAL DEFAULT 0.0,
                    Corner1Y            REAL DEFAULT 0.0,
                    Corner1Z            REAL DEFAULT 0.0,
                    Corner2X            REAL DEFAULT 0.0,
                    Corner2Y            REAL DEFAULT 0.0,
                    Corner2Z            REAL DEFAULT 0.0,
                    Corner3X            REAL DEFAULT 0.0,
                    Corner3Y            REAL DEFAULT 0.0,
                    Corner3Z            REAL DEFAULT 0.0,
                    Corner4X            REAL DEFAULT 0.0,
                    Corner4Y            REAL DEFAULT 0.0,
                    Corner4Z            REAL DEFAULT 0.0,
                    CreatedAt           DATETIME NOT NULL DEFAULT (datetime('now', '+5 hours', '+30 minutes')),
                    UpdatedAt           DATETIME NOT NULL DEFAULT (datetime('now', '+5 hours', '+30 minutes'))
                )", transaction);

            // ✅ CRITICAL FIX: Drop any legacy triggers that reference CombinedSleeves_Old
            // These triggers were created by old migration code and cause "no such table" errors
            try
            {
                using (var cmd = _connection.CreateCommand())
                {
                    cmd.Transaction = transaction;
                    
                    // Get all trigger names for CombinedSleeves table
                    cmd.CommandText = "SELECT name FROM sqlite_master WHERE type='trigger' AND tbl_name='CombinedSleeves'";
                    var triggerNames = new List<string>();
                    
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            triggerNames.Add(reader.GetString(0));
                        }
                    }
                    
                    // Drop each trigger
                    foreach (var triggerName in triggerNames)
                    {
                        ExecuteCommand($"DROP TRIGGER IF EXISTS {triggerName}", transaction);
                        _logger($"[SQLite] 🧹 Dropped legacy trigger: {triggerName}");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger($"[SQLite] ⚠️ Error dropping legacy triggers: {ex.Message}");
            }
            
            // ✅ NOTE: DeterministicGuid is now included in CREATE TABLE schema above (line 1538)
            // AddColumnIfMissing is NOT needed for fresh databases - only for migrating old DBs that lack this column
            // Wrapped in try-catch to handle edge cases where table creation might have failed
            try
            {
                AddColumnIfMissing("CombinedSleeves", "DeterministicGuid", "TEXT", transaction);
            }
            catch (Exception ex)
            {
                _logger($"[SQLite] ⚠️ Could not add DeterministicGuid column (may already exist): {ex.Message}");
            }
            
            // Table 2: Constituents (One-to-many relationship)
            ExecuteCommand(@"CREATE TABLE IF NOT EXISTS CombinedSleeveConstituents (
                    ConstituentId       INTEGER PRIMARY KEY AUTOINCREMENT,
                    CombinedSleeveId    INTEGER NOT NULL,
                    ConstituentType     TEXT NOT NULL,
                    Category            TEXT NOT NULL,
                    ClashZoneId         INTEGER,
                    ClashZoneGuid       TEXT,
                    ClusterSleeveId     INTEGER,
                    ClusterInstanceId   INTEGER,
                    CreatedAt           DATETIME NOT NULL DEFAULT (datetime('now', '+5 hours', '+30 minutes')),
                    FOREIGN KEY(CombinedSleeveId) REFERENCES CombinedSleeves(CombinedSleeveId) ON DELETE CASCADE
                )", transaction);

            // Indexes for Combined Sleeves
            ExecuteCommand("CREATE INDEX IF NOT EXISTS idx_combined_instance ON CombinedSleeves(CombinedInstanceId)", transaction);
            ExecuteCommand("CREATE INDEX IF NOT EXISTS idx_combined_guid ON CombinedSleeves(DeterministicGuid)", transaction);
            ExecuteCommand("CREATE INDEX IF NOT EXISTS idx_combined_combo ON CombinedSleeves(ComboId)", transaction);
            ExecuteCommand("CREATE INDEX IF NOT EXISTS idx_combined_filter ON CombinedSleeves(FilterId)", transaction);

            // Indexes for Constituents
            ExecuteCommand("CREATE INDEX IF NOT EXISTS idx_const_combined_id ON CombinedSleeveConstituents(CombinedSleeveId)", transaction);
            ExecuteCommand("CREATE INDEX IF NOT EXISTS idx_const_clash_zone ON CombinedSleeveConstituents(ClashZoneId)", transaction);
            ExecuteCommand("CREATE INDEX IF NOT EXISTS idx_const_cluster_id ON CombinedSleeveConstituents(ClusterSleeveId)", transaction);
        }

        /// <summary>
        /// ✅ SESSION CONTEXT: Create SessionContext table for storing session-level data
        /// Stores section box bounds and other session data for "dump once, use many times" pattern
        /// Simple key-value store for one-time session data (not bulk operations)
        /// </summary>
        private void EnsureSessionContextTable(SQLiteTransaction transaction)
        {
            ExecuteCommand(@"
                CREATE TABLE IF NOT EXISTS SessionContext (
                    Key TEXT PRIMARY KEY,
                    Value TEXT,
                    UpdatedAt DATETIME NOT NULL DEFAULT (datetime('now', '+5 hours', '+30 minutes'))
                )", transaction);

            _logger("[SQLite] ✅ SessionContext table ready (stores section box bounds and session data)");
        }

        /// <summary>
        /// Clears all data from sleeve-related tables.
        /// Used by the ClearAllSleeveDbTablesCommand for resetting the database.
        /// </summary>
        public void ClearAllTables()
        {
#if NET8_0_OR_GREATER
            // ✅ CRITICAL: In Microsoft.Data.Sqlite (.NET 8), PRAGMA foreign_keys cannot be toggled INSIDE a transaction.
            using (var fkOffCmd = _connection.CreateCommand())
            {
                fkOffCmd.CommandText = "PRAGMA foreign_keys = OFF;";
                fkOffCmd.ExecuteNonQuery();
            }
#endif

            using (var transaction = _connection.BeginTransaction())
            {
                try
                {
#if !NET8_0_OR_GREATER
                    // Legacy behavior for System.Data.SQLite
                    ExecuteCommand("PRAGMA foreign_keys = OFF;", transaction);
#endif

                    // Clear tables in dependency order (reverse creation order roughly)
                    ExecuteCommand("DELETE FROM SleeveEvents;", transaction);
                    ExecuteCommand("DELETE FROM ClashZones;", transaction);
                    ExecuteCommand("DELETE FROM CombinedSleeveConstituents;", transaction);
                    ExecuteCommand("DELETE FROM CombinedSleeves;", transaction);
                    ExecuteCommand("DELETE FROM ClusterSleeves;", transaction);
                    ExecuteCommand("DELETE FROM ClusterSleeves_v2;", transaction); // ✅ BATCH V2: Clear cluster calculation data
                    ExecuteCommand("DELETE FROM SleeveSnapshots;", transaction);
                    ExecuteCommand("DELETE FROM ParameterTransferFlags;", transaction);
                    ExecuteCommand("DELETE FROM CategoryProcessingMarkers;", transaction);
                    ExecuteCommand("DELETE FROM Conditions;", transaction);
                    ExecuteCommand("DELETE FROM FileCombos;", transaction);
                    ExecuteCommand("DELETE FROM Filters;", transaction);
                    
                    // Reset auto-increment counters
                    ExecuteCommand("DELETE FROM sqlite_sequence WHERE name IN ('SleeveEvents', 'ClashZones', 'ClusterSleeves', 'ClusterSleeves_v2', 'SleeveSnapshots', 'ParameterTransferFlags', 'CategoryProcessingMarkers', 'Conditions', 'FileCombos', 'Filters');", transaction);

#if !NET8_0_OR_GREATER
                    // Re-enable foreign keys inside transaction for System.Data.SQLite
                    ExecuteCommand("PRAGMA foreign_keys = ON;", transaction);
#endif

                    transaction.Commit();
                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    _logger($"[SQLite] ❌ Error clearing tables: {ex.Message}");
                    throw;
                }
            }

#if NET8_0_OR_GREATER
            // Re-enable foreign keys OUTSIDE the transaction for .NET 8
            using (var fkOnCmd = _connection.CreateCommand())
            {
                fkOnCmd.CommandText = "PRAGMA foreign_keys = ON;";
                fkOnCmd.ExecuteNonQuery();
            }
#endif
            _logger("[SQLite] ✅ All tables cleared successfully.");
        }
    }
}

