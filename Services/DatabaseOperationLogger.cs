using System;
using System.Collections.Generic;
using System.Text;
using JSE_RevitAddin_MEP_OPENINGS.Data;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// ✅ COMPREHENSIVE DATABASE LOGGING SERVICE
    /// Logs all database operations (INSERT, UPDATE, SELECT, DELETE) with table names, columns, and values
    /// Used to verify database operations match flowchart expectations
    /// </summary>
    public class DatabaseOperationLogger
    {
        private static readonly object _lock = new object();

        /// <summary>
        /// Log a database operation with full details
        /// </summary>
        public static void LogOperation(
            string operation, // INSERT, UPDATE, SELECT, DELETE
            string tableName,
            Dictionary<string, object> parameters = null,
            int rowsAffected = -1,
            string additionalInfo = null)
        {
            if (DeploymentConfiguration.DeploymentMode)
                return;

            lock (_lock)
            {
                try
                {
                    var logEntry = new StringBuilder();
                    logEntry.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] === DATABASE OPERATION ===");
                    logEntry.AppendLine($"Operation: {operation}");
                    logEntry.AppendLine($"Table: {tableName}");

                    if (parameters != null && parameters.Count > 0)
                    {
                        logEntry.AppendLine("Parameters:");
                        foreach (var param in parameters)
                        {
                            var value = param.Value?.ToString() ?? "NULL";
                            if (value.Length > 100)
                                value = value.Substring(0, 100) + "...";
                            logEntry.AppendLine($"  {param.Key} = {value}");
                        }
                    }

                    if (rowsAffected >= 0)
                        logEntry.AppendLine($"Rows Affected: {rowsAffected}");

                    if (!string.IsNullOrWhiteSpace(additionalInfo))
                        logEntry.AppendLine($"Info: {additionalInfo}");

                    logEntry.AppendLine("---");
                    logEntry.AppendLine();

                    SafeFileLogger.SafeAppendText("database_operations.log", logEntry.ToString());
                }
                catch
                {
                    // Silent fail - don't break operations if logging fails
                }
            }
        }

        /// <summary>
        /// Log a SELECT query with results
        /// </summary>
        public static void LogSelect(
            string tableName,
            string whereClause = null,
            int resultCount = -1,
            Dictionary<string, object> sampleRow = null,
            string additionalInfo = null)
        {
            if (DeploymentConfiguration.DeploymentMode)
                return;

            lock (_lock)
            {
                try
                {
                    var logEntry = new StringBuilder();
                    logEntry.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] === DATABASE SELECT ===");
                    logEntry.AppendLine($"Table: {tableName}");

                    if (!string.IsNullOrWhiteSpace(whereClause))
                        logEntry.AppendLine($"WHERE: {whereClause}");

                    if (resultCount >= 0)
                        logEntry.AppendLine($"Results: {resultCount} row(s)");

                    if (sampleRow != null && sampleRow.Count > 0)
                    {
                        logEntry.AppendLine("Sample Row:");
                        foreach (var col in sampleRow)
                        {
                            var value = col.Value?.ToString() ?? "NULL";
                            if (value.Length > 100)
                                value = value.Substring(0, 100) + "...";
                            logEntry.AppendLine($"  {col.Key} = {value}");
                        }
                    }

                    if (!string.IsNullOrWhiteSpace(additionalInfo))
                        logEntry.AppendLine($"Info: {additionalInfo}");

                    logEntry.AppendLine("---");
                    logEntry.AppendLine();

                    SafeFileLogger.SafeAppendText("database_operations.log", logEntry.ToString());
                }
                catch
                {
                    // Silent fail
                }
            }
        }

        /// <summary>
        /// Log transaction operations
        /// </summary>
        public static void LogTransaction(string operation, string status, string errorMessage = null)
        {
            if (DeploymentConfiguration.DeploymentMode)
                return;

            lock (_lock)
            {
                try
                {
                    var logEntry = new StringBuilder();
                    logEntry.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] === TRANSACTION ===");
                    logEntry.AppendLine($"Operation: {operation}");
                    logEntry.AppendLine($"Status: {status}");

                    if (!string.IsNullOrWhiteSpace(errorMessage))
                        logEntry.AppendLine($"Error: {errorMessage}");

                    logEntry.AppendLine("---");
                    logEntry.AppendLine();

                    SafeFileLogger.SafeAppendText("database_operations.log", logEntry.ToString());
                }
                catch
                {
                    // Silent fail
                }
            }
        }

        /// <summary>
        /// Log table schema verification
        /// </summary>
        public static void LogTableSchema(string tableName, List<string> columns)
        {
            if (DeploymentConfiguration.DeploymentMode)
                return;

            lock (_lock)
            {
                try
                {
                    var logEntry = new StringBuilder();
                    logEntry.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] === TABLE SCHEMA ===");
                    logEntry.AppendLine($"Table: {tableName}");
                    logEntry.AppendLine($"Columns ({columns.Count}): {string.Join(", ", columns)}");
                    logEntry.AppendLine("---");
                    logEntry.AppendLine();

                    SafeFileLogger.SafeAppendText("database_operations.log", logEntry.ToString());
                }
                catch
                {
                    // Silent fail
                }
            }
        }
    }
}


