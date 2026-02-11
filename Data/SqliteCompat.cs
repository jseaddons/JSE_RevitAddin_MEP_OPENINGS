// ✅ SQLite Provider Compatibility Layer
// R23/R24 (net48): Uses System.Data.SQLite — original using directives in each file.
// R25/R26 (net8.0): Uses Microsoft.Data.Sqlite — global aliases below map
//   the old PascalCase names (SQLiteConnection) to the new camelCase names (SqliteConnection),
//   so all existing code compiles unchanged against the faster .NET 8 provider.
//
// NOTE: SQLiteParameter is NOT a global alias — it uses a wrapper class (below)
// because Microsoft.Data.Sqlite.SqliteParameter(string, SqliteType) constructor
// is incompatible with System.Data.SQLite.SQLiteParameter(string, DbType).

#if NET8_0_OR_GREATER
global using SQLiteConnection = Microsoft.Data.Sqlite.SqliteConnection;
global using SQLiteCommand = Microsoft.Data.Sqlite.SqliteCommand;
global using SQLiteTransaction = Microsoft.Data.Sqlite.SqliteTransaction;
global using SQLiteDataReader = Microsoft.Data.Sqlite.SqliteDataReader;
// SQLiteParameter: see wrapper class below (not aliased due to constructor difference)
global using SQLiteException = Microsoft.Data.Sqlite.SqliteException;
#endif

/// <summary>
/// Connection string helper: Version=3 is required by System.Data.SQLite but not supported by Microsoft.Data.Sqlite.
/// </summary>
internal static class SqliteConnStr
{
    internal static string Build(string dbPath) =>
#if NET8_0_OR_GREATER
        $"Data Source={dbPath}";
#else
        $"Data Source={dbPath};Version=3;";
#endif
}

#if NET8_0_OR_GREATER
/// <summary>
/// Compatibility wrapper for SQLiteParameter on NET8+.
/// System.Data.SQLite has SQLiteParameter(string, DbType) but Microsoft.Data.Sqlite
/// has SqliteParameter(string, SqliteType) — different enum type in 2nd arg.
/// This wrapper provides both constructor signatures so all existing code compiles unchanged.
/// </summary>
internal class SQLiteParameter : Microsoft.Data.Sqlite.SqliteParameter
{
    public SQLiteParameter() : base() { }
    public SQLiteParameter(string parameterName, object value) : base(parameterName, value) { }
    public SQLiteParameter(string parameterName, System.Data.DbType dbType) : base()
    {
        ParameterName = parameterName;
        DbType = dbType;
    }
}

/// <summary>
/// Extension method for SqliteParameterCollection.Add(string, DbType).
/// System.Data.SQLite has this overload but Microsoft.Data.Sqlite only has Add(string, SqliteType).
/// The compiler will use this extension when no matching instance method is found.
/// </summary>
internal static class SqliteParameterCollectionExtensions
{
    internal static SQLiteParameter Add(
        this Microsoft.Data.Sqlite.SqliteParameterCollection collection,
        string parameterName,
        System.Data.DbType dbType)
    {
        var param = new SQLiteParameter(parameterName, dbType);
        collection.Add(param);
        return param;
    }
}
#endif
