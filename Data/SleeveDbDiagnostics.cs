using System;
using System.Data;
using System.IO;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services;
#if NET8_0_OR_GREATER
using Microsoft.Data.Sqlite;
#else
using System.Data.SQLite;
#endif

namespace JSE_RevitAddin_MEP_OPENINGS.Data
{
    /// <summary>
    /// Simple database diagnostics - tests connection and prompts user on failure.
    /// Does NOT fallback to XML - operation is cancelled if DB fails.
    /// </summary>
    public static class SleeveDbDiagnostics
    {
        /// <summary>
        /// Tests database connectivity and shows user-friendly error if it fails.
        /// Returns true if OK, throws exception if fails (after showing prompt).
        /// </summary>
        public static void VerifyConnectionOrPrompt(Document doc)
        {
            var filtersDir = ProjectPathService.GetFiltersDirectory(doc);
            var projectName = doc.Title ?? "Default";
            var safeName = System.Text.RegularExpressions.Regex.Replace(
                projectName, @"[^\w\s-]", string.Empty).Replace(" ", "_");
            var dbPath = Path.Combine(filtersDir, $"{safeName}_SleevePersistence.db");
            
            string errorDetails = null;
            
            // Test 1: SQLite DLLs present
            var dllCheck = CheckSQLiteDlls();
            if (!dllCheck.IsValid)
            {
                errorDetails = $"SQLite libraries missing:\n{dllCheck.ErrorMessage}\n\n" +
                    "Required files:\n" +
#if NET8_0_OR_GREATER
                    "• e_sqlite3.dll\n\n" +
#else
                    "• System.Data.SQLite.dll\n" +
                    "• x64\\SQLite.Interop.dll\n\n" +
#endif
                    "To fix:\n" +
                    "1. Close ALL Revit instances\n" +
                    "2. Run deploy-addin.ps1\n" +
                    "3. Restart Revit";
                    
                ShowErrorAndThrow("SQLite Libraries Missing", errorDetails);
            }
            
            // Test 2: Can open database connection
            try
            {
                using var testConnection = CreateTestConnection(dbPath);
                testConnection.Open();
                
                // Test a simple query
                using var cmd = testConnection.CreateCommand();
                cmd.CommandText = "SELECT 1";
                cmd.ExecuteScalar();
                
                // All good - return silently
            }
            catch (Exception ex)
            {
                errorDetails = $"Cannot open database:\n{dbPath}\n\n" +
                    "Error: " + ex.Message + "\n\n" +
                    "Possible causes:\n" +
                    "• Database file is locked by another Revit instance\n" +
                    "• Folder permissions issue\n" +
                    "• Database file is corrupted\n\n" +
                    "To fix:\n" +
                    "1. Close ALL Revit instances and retry\n" +
                    "2. Check if .db file is read-only\n" +
                    "3. Delete the .db file to recreate it\n\n" +
                    "Contact support if issue persists.";
                    
                ShowErrorAndThrow("Database Connection Failed", errorDetails);
            }
        }
        
        /// <summary>
        /// Shows error dialog and throws exception to stop operation
        /// </summary>
        private static void ShowErrorAndThrow(string title, string message)
        {
            MessageBox.Show(
                message,
                $"❌ {title}",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
                
            throw new InvalidOperationException($"{title}: {message}");
        }
        
        /// <summary>
        /// Checks if required SQLite DLLs are present
        /// </summary>
        private static SQLiteCheckResult CheckSQLiteDlls()
        {
            var assemblyDir = Path.GetDirectoryName(
                System.Reflection.Assembly.GetExecutingAssembly().Location);
                
#if NET8_0_OR_GREATER
            var esqlite = Path.Combine(assemblyDir, "e_sqlite3.dll");
            if (!File.Exists(esqlite))
            {
                return new SQLiteCheckResult
                {
                    IsValid = false,
                    ErrorMessage = "Missing e_sqlite3.dll"
                };
            }
#else
            var sqlite = Path.Combine(assemblyDir, "System.Data.SQLite.dll");
            var interop = Path.Combine(assemblyDir, "x64", "SQLite.Interop.dll");
            
            if (!File.Exists(sqlite))
            {
                return new SQLiteCheckResult { IsValid = false, ErrorMessage = "Missing System.Data.SQLite.dll" };
            }
            if (!File.Exists(interop))
            {
                return new SQLiteCheckResult { IsValid = false, ErrorMessage = "Missing x64\\SQLite.Interop.dll" };
            }
#endif
            return new SQLiteCheckResult { IsValid = true };
        }
        
        private static IDbConnection CreateTestConnection(string dbPath)
        {
#if NET8_0_OR_GREATER
            return new SqliteConnection($"Data Source={dbPath}");
#else
            var builder = new SQLiteConnectionStringBuilder
            {
                DataSource = dbPath,
                ForeignKeys = true,
                JournalMode = SQLiteJournalModeEnum.Wal
            };
            return new SQLiteConnection(builder.ConnectionString);
#endif
        }
        
        private class SQLiteCheckResult
        {
            public bool IsValid { get; set; }
            public string ErrorMessage { get; set; } = string.Empty;
        }
    }
}
