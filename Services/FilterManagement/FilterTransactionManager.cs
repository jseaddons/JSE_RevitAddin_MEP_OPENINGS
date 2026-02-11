using System;
#if !NET8_0_OR_GREATER
using System.Data.SQLite;
#endif
using System.Threading;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor;
using JSE_RevitAddin_MEP_OPENINGS.Services.Logging;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.FilterManagement
{
    /// <summary>
    /// Team J: Transaction manager for filter operations.
    /// Provides atomic operations with retry logic and fail-safe rollback.
    /// 
    /// ✅ FEATURES:
    /// - Transaction management (atomic operations)
    /// - Retry logic for database busy/locked errors
    /// - Fail-safe rollback on errors
    /// - Connection reuse within transaction scope
    /// </summary>
    public class FilterTransactionManager
    {
        private readonly Document _document;
        private readonly ILogger _logger;
        private const int MaxRetries = 3;
        private const int BaseRetryDelayMs = 100;

        public FilterTransactionManager(Document document, ILogger logger = null)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _logger = logger ?? LoggerAdapter.Default;
        }

        /// <summary>
        /// Executes an operation within a transaction with retry logic.
        /// ✅ ATOMIC: All operations succeed or all fail (rollback).
        /// ✅ RETRY: Automatically retries on database busy/locked errors.
        /// ✅ FAIL-SAFE: Automatic rollback on any error.
        /// </summary>
        public TResult ExecuteInTransaction<TResult>(
            Func<SleeveDbContext, SQLiteTransaction, TResult> operation,
            string operationName)
        {
            if (operation == null)
                throw new ArgumentNullException(nameof(operation));

            for (int attempt = 1; attempt <= MaxRetries; attempt++)
            {
                try
                {
                    using (var context = new SleeveDbContext(_document, msg => _logger.Info($"[FilterTransaction] {msg}")))
                    {
                        using (var transaction = context.Connection.BeginTransaction())
                        {
                            try
                            {
                                _logger.Info($"Starting transaction: {operationName} (attempt {attempt}/{MaxRetries})", "FilterTransactionManager");
                                
                                var result = operation(context, transaction);
                                
                                transaction.Commit();
                                _logger.Info($"Transaction committed: {operationName}", "FilterTransactionManager");
                                
                                return result;
                            }
                            catch (Exception ex)
                            {
                                transaction.Rollback();
                                _logger.Error($"Transaction rolled back: {operationName} - {ex.Message}", ex, "FilterTransactionManager");
                                
                                // Check if this is a retryable error
                                if (IsRetryableError(ex) && attempt < MaxRetries)
                                {
                                    var delayMs = BaseRetryDelayMs * (int)Math.Pow(2, attempt - 1); // Exponential backoff
                                    _logger.Warning($"Retrying {operationName} after {delayMs}ms (attempt {attempt}/{MaxRetries})", "FilterTransactionManager");
                                    Thread.Sleep(delayMs);
                                    continue; // Retry
                                }
                                
                                throw; // Re-throw if not retryable or max retries reached
                            }
                        }
                    }
                }
                catch (SQLiteException sqlEx) when (IsRetryableError(sqlEx) && attempt < MaxRetries)
                {
                    var delayMs = BaseRetryDelayMs * (int)Math.Pow(2, attempt - 1);
                    _logger.Warning($"Retrying {operationName} after {delayMs}ms (attempt {attempt}/{MaxRetries})", "FilterTransactionManager");
                    Thread.Sleep(delayMs);
                    continue; // Retry
                }
            }

            throw new InvalidOperationException($"Failed to execute {operationName} after {MaxRetries} attempts");
        }

        /// <summary>
        /// Executes an operation within a transaction (void return).
        /// </summary>
        public void ExecuteInTransaction(
            Action<SleeveDbContext, SQLiteTransaction> operation,
            string operationName)
        {
            ExecuteInTransaction<object>((context, transaction) =>
            {
                operation(context, transaction);
                return null;
            }, operationName);
        }

        /// <summary>
        /// Checks if an error is retryable (database busy/locked).
        /// </summary>
        private bool IsRetryableError(Exception ex)
        {
#if NET8_0_OR_GREATER
            // Microsoft.Data.Sqlite: SqliteErrorCode is int (SQLITE_BUSY=5, SQLITE_LOCKED=6)
            if (ex is SQLiteException sqlEx)
            {
                return sqlEx.SqliteErrorCode == 5 || sqlEx.SqliteErrorCode == 6;
            }
            if (ex.InnerException is SQLiteException innerSqlEx)
            {
                return innerSqlEx.SqliteErrorCode == 5 || innerSqlEx.SqliteErrorCode == 6;
            }
#else
            if (ex is SQLiteException sqlEx)
            {
                return sqlEx.ResultCode == SQLiteErrorCode.Busy ||
                       sqlEx.ResultCode == SQLiteErrorCode.Locked;
            }
            if (ex.InnerException is SQLiteException innerSqlEx)
            {
                return innerSqlEx.ResultCode == SQLiteErrorCode.Busy ||
                       innerSqlEx.ResultCode == SQLiteErrorCode.Locked;
            }
#endif
            return false;
        }
    }
}

