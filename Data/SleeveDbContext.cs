using System;
using System.Collections.Generic;
using System.Data.SQLite;
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
                VerifyDependency("System.Data.SQLite.dll", assemblyDirectory);
                VerifyDependency(Path.Combine("x64", "SQLite.Interop.dll"), assemblyDirectory);

                var filtersDirectory = ProjectPathService.GetFiltersDirectory(document);
                Directory.CreateDirectory(filtersDirectory);
                _logger($"[SQLite] Filters directory: {filtersDirectory}");

                var projectName = document.Title ?? "Default";
                var safeProjectName = Regex.Replace(projectName, @"[^\w\s-]", string.Empty).Replace(" ", "_");
                _databasePath = Path.Combine(filtersDirectory, $"{safeProjectName}_SleevePersistence.db");
                _logger($"[SQLite] Database path: {_databasePath}");

                var builder = new SQLiteConnectionStringBuilder
                {
                    DataSource = _databasePath,
                    ForeignKeys = true,
                    JournalMode = SQLiteJournalModeEnum.Wal
                };

                _connection = new SQLiteConnection(builder.ConnectionString);
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

                EnsureSchemaCreated();
                EnsureSchemaUpgraded();
                
                // ✅ R-TREE CHECK: Test if R-tree extension is available
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
                throw;
            }
        }

        private void VerifyDependency(string relativePath, string assemblyDirectory)
        {
            var targetPath = Path.Combine(assemblyDirectory, relativePath);
            if (File.Exists(targetPath))
            {
                _logger($"[SQLite] ✅ Found dependency: {relativePath}");
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
                    return;
                }
                catch (Exception copyEx)
                {
                    _logger($"[SQLite] ❌ Failed to copy dependency from {candidate}: {copyEx.Message}");
                }
            }

            _logger($"[SQLite] ⚠️ Dependency {relativePath} remains missing. SQLite may fail to load.");
        }

        private static IEnumerable<string> FindInNuGet(string fileName)
        {
            var nugetRoot = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            if (string.IsNullOrEmpty(nugetRoot))
            {
                yield break;
            }

            var sqlitePackageRoot = Path.Combine(nugetRoot, @".nuget\packages\system.data.sqlite.core");
            if (!Directory.Exists(sqlitePackageRoot))
            {
                yield break;
            }

            foreach (var candidate in Directory.EnumerateFiles(sqlitePackageRoot, fileName, SearchOption.AllDirectories))
            {
                if (fileName.Equals("SQLite.Interop.dll", StringComparison.OrdinalIgnoreCase) &&
                    !candidate.Contains(@"x64", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
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
                            SleeveInstanceId INTEGER,
                            ClusterInstanceId INTEGER,
                            SleeveWidth   REAL,
                            SleeveHeight  REAL,
                            SleeveDiameter REAL,
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
                            ConditionId   INTEGER PRIMARY KEY AUTOINCREMENT,
                            FilterId      INTEGER NOT NULL,
                            Category      TEXT NOT NULL,
                            RectNormal    REAL,
                            RectInsulated REAL,
                            RoundNormal   REAL,
                            RoundInsulated REAL,
                            PipesNormal   REAL,
                            PipesInsulated REAL,
                            CableTrayTop  REAL,
                            CableTrayOther REAL,
                            OpeningPrefs  TEXT,
                            UpdatedAt     DATETIME NOT NULL DEFAULT (datetime('now', '+5 hours', '+30 minutes')),
                            FOREIGN KEY(FilterId) REFERENCES Filters(FilterId) ON DELETE CASCADE,
                            UNIQUE(FilterId, Category)
                        )", transaction);

                    EnsureSleeveSnapshotTable(transaction);
                    EnsureClusterSleevesTable(transaction);
                    
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
                    EnsureClusterSleevesTable(transaction);
                    
                    // ✅ R-TREE: Create R-tree table for existing databases (if enabled)
                    EnsureRTreeTable(transaction);

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
                    // ✅ PIPE DIAMETER COLUMNS: Add outer diameter and nominal diameter columns for pipes
                    if (AddColumnIfMissing("ClashZones", "MepElementOuterDiameter", "REAL DEFAULT 0.0", transaction))
                        _logger("[SQLite] ✅ Added MepElementOuterDiameter column to ClashZones (pipe outer diameter in feet)");
                    if (AddColumnIfMissing("ClashZones", "MepElementNominalDiameter", "REAL DEFAULT 0.0", transaction))
                        _logger("[SQLite] ✅ Added MepElementNominalDiameter column to ClashZones (pipe nominal diameter in feet)");
                    // ✅ SIZE PARAMETER VALUE: Add string column for Size parameter value (e.g., "20 mmø", "200 mm dia symbol")
                    // This is the exact text from the Size parameter, stored for transfer to sleeve MEP_Size parameter
                    if (AddColumnIfMissing("ClashZones", "MepElementSizeParameterValue", "TEXT DEFAULT ''", transaction))
                        _logger("[SQLite] ✅ Added MepElementSizeParameterValue column to ClashZones (Size parameter as string)");
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
                    AddColumnIfMissing("ClashZones", "IsClusteredFlag", "INTEGER", transaction);
                    AddColumnIfMissing("ClashZones", "MarkedForClusterProcess", "INTEGER", transaction);
                    AddColumnIfMissing("ClashZones", "AfterClusterSleeveId", "INTEGER", transaction);
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
                    AddColumnIfMissing("Conditions", "DuctAccessoryMepNormal", "REAL", transaction);
                    AddColumnIfMissing("Conditions", "DuctAccessoryOtherNormal", "REAL", transaction);
                    AddColumnIfMissing("Conditions", "OpeningPrefs", "TEXT", transaction);
                    AddColumnIfMissing("Conditions", "HorizontalLevel", "TEXT", transaction);
                    AddColumnIfMissing("Conditions", "VerticalLevel", "TEXT", transaction);
                    AddColumnIfMissing("Conditions", "CreationMode", "TEXT", transaction);

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

                    // ✅ OPTION 4 IMPLEMENTATION: Add triggers, constraints, and views for sophisticated flag management
                    // ✅ TEMPORARILY DISABLED: Triggers causing SQL logic errors - will re-enable with compatible syntax
                    // EnsureFlagManagementTriggers(transaction);
                    EnsureFlagManagementConstraints(transaction);
                    
                    // ✅ TEMPORARILY DISABLED: Views causing SQL logic errors - will re-enable after fixing column references
                    // EnsureFlagManagementViews(transaction);
                    
                    // ✅ DATABASE GUID MANAGEMENT: Add views for GUID lookups
                    // ✅ TEMPORARILY DISABLED: Views causing SQL logic errors
                    // EnsureGuidManagementViews(transaction);

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
            try
            {
                ExecuteCommand(@"
                    ALTER TABLE ClusterSleeves ADD COLUMN ClashZoneGuids TEXT;
                ", transaction);
            }
            catch
            {
                // Column already exists, ignore
            }
            
            try
            {
                ExecuteCommand(@"
                    ALTER TABLE ClusterSleeves ADD COLUMN MepSizes TEXT;
                ", transaction);
            }
            catch
            {
                // Column already exists, ignore
            }
            
            try
            {
                ExecuteCommand(@"
                    ALTER TABLE ClusterSleeves ADD COLUMN MepSystemNames TEXT;
                ", transaction);
            }
            catch
            {
                // Column already exists, ignore
            }
            
            try
            {
                ExecuteCommand(@"
                    ALTER TABLE ClusterSleeves ADD COLUMN MepElementIds TEXT;
                ", transaction);
            }
            catch
            {
                // Column already exists, ignore
            }
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
                _connection?.Close();
                _connection?.Dispose();
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
    }
}

