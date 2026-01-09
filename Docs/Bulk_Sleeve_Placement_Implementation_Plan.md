# Sleeve Placement Implementation Plan

**Document Version:** 1.0  
**Date:** December 2025  
**Status:** Implementation Guide

---

## Table of Contents

1. [Overview](#1-overview)
2. [Implementation Phases](#2-implementation-phases)
3. [Phase 1: Error Handling Implementation](#3-phase-1-error-handling-implementation)
4. [Phase 2: Database Operations Enhancement](#4-phase-2-database-operations-enhancement)
5. [Phase 3: Transaction Management](#5-phase-3-transaction-management)
6. [Phase 4: Logging and Monitoring](#6-phase-4-logging-and-monitoring)
7. [Phase 5: User Experience Improvements](#7-phase-5-user-experience-improvements)
8. [Testing Strategy](#8-testing-strategy)
9. [Deployment Checklist](#9-deployment-checklist)

---

## 1. Overview

This implementation plan provides a step-by-step guide for implementing comprehensive error handling and resilience features in the sleeve placement system. The plan is organized into phases that can be implemented incrementally.

### 1.1 Goals

- ✅ **Robust Error Handling** - Handle all error scenarios gracefully
- ✅ **Data Integrity** - Prevent partial data corruption through transactions
- ✅ **User Experience** - Clear error messages and progress reporting
- ✅ **Maintainability** - Comprehensive logging for debugging
- ✅ **Resilience** - System continues operation where possible

### 1.2 Prerequisites

- ✅ Refactored refresh service in place
- ✅ Database-first architecture implemented
- ✅ Repository pattern established
- ✅ Transaction infrastructure available

---

## 2. Implementation Phases

| Phase | Focus | Duration | Priority |
|-------|-------|----------|----------|
| **Phase 1** | Error Handling Infrastructure | 2-3 days | High |
| **Phase 2** | Database Operations Enhancement | 2-3 days | High |
| **Phase 3** | Transaction Management | 1-2 days | High |
| **Phase 4** | Logging and Monitoring | 1-2 days | Medium |
| **Phase 5** | User Experience Improvements | 1-2 days | Medium |

**Total Estimated Duration:** 7-12 days

---

## 3. Phase 1: Error Handling Implementation (OOP-Based)

### 3.1 Custom Exception Classes

**Files to Create:**
- `Exceptions/DatabaseException.cs`
- `Exceptions/RepositoryException.cs`
- `Exceptions/TransactionException.cs`
- `Exceptions/ValidationException.cs`
- `Exceptions/SleevePlacementException.cs`

**Implementation Steps:**

1. **Create Base Exception Class:**
```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Exceptions
{
    public abstract class SleevePlacementBaseException : Exception
    {
        public string OperationName { get; }
        public Dictionary<string, object> Context { get; }
        public DateTime Timestamp { get; }

        protected SleevePlacementBaseException(
            string operationName,
            string message,
            Exception innerException = null,
            Dictionary<string, object> context = null)
            : base(message, innerException)
        {
            OperationName = operationName;
            Context = context ?? new Dictionary<string, object>();
            Timestamp = DateTime.Now;
        }

        public abstract string GetUserFriendlyMessage();
    }
}
```

2. **Create Specific Exception Classes:**
```csharp
public class DatabaseException : SleevePlacementBaseException
{
    public SQLiteErrorCode? ErrorCode { get; }

    public DatabaseException(
        string operationName,
        SQLiteException innerException,
        Dictionary<string, object> context = null)
        : base(operationName, $"Database operation failed: {innerException.Message}", innerException, context)
    {
        ErrorCode = innerException.ResultCode;
    }

    public override string GetUserFriendlyMessage()
    {
        return ErrorCode switch
        {
            SQLiteErrorCode.Corrupt => "The database file appears to be corrupted. Please restore from backup.",
            SQLiteErrorCode.Locked => "The database is currently in use. Please close other instances and try again.",
            SQLiteErrorCode.Constraint => "A duplicate entry was detected. Please check your data and try again.",
            SQLiteErrorCode.Busy => "The database is busy. Please wait a moment and try again.",
            _ => "A database error occurred. Please contact support with the error details."
        };
    }
}

public class RepositoryException : SleevePlacementBaseException
{
    public string RepositoryName { get; }
    public string EntityType { get; }

    public RepositoryException(
        string repositoryName,
        string entityType,
        string operationName,
        Exception innerException,
        Dictionary<string, object> context = null)
        : base(operationName, $"Repository operation failed in {repositoryName}", innerException, context)
    {
        RepositoryName = repositoryName;
        EntityType = entityType;
    }

    public override string GetUserFriendlyMessage()
    {
        return $"Failed to {OperationName} {EntityType}. Please try again or contact support.";
    }
}

public class TransactionException : SleevePlacementBaseException
{
    public bool WasRolledBack { get; }

    public TransactionException(
        string operationName,
        Exception innerException,
        bool wasRolledBack = true,
        Dictionary<string, object> context = null)
        : base(operationName, "Transaction failed and was rolled back", innerException, context)
    {
        WasRolledBack = wasRolledBack;
    }

    public override string GetUserFriendlyMessage()
    {
        return "The operation could not be completed. All changes have been rolled back. Please try again.";
    }
}

public class ValidationException : SleevePlacementBaseException
{
    public List<string> ValidationErrors { get; }

    public ValidationException(
        string operationName,
        List<string> validationErrors,
        Dictionary<string, object> context = null)
        : base(operationName, $"Validation failed: {string.Join(", ", validationErrors)}", null, context)
    {
        ValidationErrors = validationErrors;
    }

    public override string GetUserFriendlyMessage()
    {
        return $"Validation failed:\n{string.Join("\n", ValidationErrors)}";
    }
}

public class SleevePlacementException : SleevePlacementBaseException
{
    public string ClashZoneId { get; }
    public string Category { get; }

    public SleevePlacementException(
        string clashZoneId,
        string category,
        string operationName,
        Exception innerException,
        Dictionary<string, object> context = null)
        : base(operationName, $"Sleeve placement failed for zone {clashZoneId}", innerException, context)
    {
        ClashZoneId = clashZoneId;
        Category = category;
    }

    public override string GetUserFriendlyMessage()
    {
        return $"Failed to place sleeve for {Category}. The operation will continue with other sleeves.";
    }
}
```

### 3.2 Error Handler Interface and Implementations

**Files to Create:**
- `Services/ErrorHandling/IErrorHandler.cs`
- `Services/ErrorHandling/DatabaseErrorHandler.cs`
- `Services/ErrorHandling/RepositoryErrorHandler.cs`
- `Services/ErrorHandling/TransactionErrorHandler.cs`
- `Services/ErrorHandling/SleevePlacementErrorHandler.cs`

**Implementation Steps:**

1. **Create Error Handler Interface:**
```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Services.ErrorHandling
{
    public interface IErrorHandler
    {
        bool CanHandle(Exception exception);
        ErrorHandlingResult Handle(Exception exception, Dictionary<string, object> context = null);
    }

    public class ErrorHandlingResult
    {
        public bool ShouldCancel { get; set; }
        public bool ShouldRetry { get; set; }
        public string UserMessage { get; set; }
        public bool ShouldLog { get; set; } = true;
        public Exception TransformedException { get; set; }
    }
}
```

2. **Implement Database Error Handler:**
```csharp
public class DatabaseErrorHandler : IErrorHandler
{
    private readonly ILoggingService _logger;

    public DatabaseErrorHandler(ILoggingService logger)
    {
        _logger = logger;
    }

    public bool CanHandle(Exception exception)
    {
        return exception is SQLiteException || 
               exception is DatabaseException ||
               (exception.InnerException is SQLiteException);
    }

    public ErrorHandlingResult Handle(Exception exception, Dictionary<string, object> context = null)
    {
        var sqliteEx = exception as SQLiteException ?? 
                      exception.InnerException as SQLiteException;

        if (sqliteEx == null)
            return new ErrorHandlingResult { ShouldCancel = true };

        var dbException = new DatabaseException(
            context?.GetValueOrDefault("OperationName")?.ToString() ?? "Database Operation",
            sqliteEx,
            context);

        _logger.LogError(dbException);

        return new ErrorHandlingResult
        {
            ShouldCancel = true,
            UserMessage = dbException.GetUserFriendlyMessage(),
            TransformedException = dbException
        };
    }
}
```

3. **Implement Repository Error Handler:**
```csharp
public class RepositoryErrorHandler : IErrorHandler
{
    private readonly ILoggingService _logger;

    public RepositoryErrorHandler(ILoggingService logger)
    {
        _logger = logger;
    }

    public bool CanHandle(Exception exception)
    {
        return exception is RepositoryException;
    }

    public ErrorHandlingResult Handle(Exception exception, Dictionary<string, object> context = null)
    {
        var repoEx = exception as RepositoryException;
        if (repoEx == null) return null;

        _logger.LogError(repoEx);

        return new ErrorHandlingResult
        {
            ShouldCancel = true,
            UserMessage = repoEx.GetUserFriendlyMessage(),
            TransformedException = repoEx
        };
    }
}
```

4. **Implement Transaction Error Handler:**
```csharp
public class TransactionErrorHandler : IErrorHandler
{
    private readonly ILoggingService _logger;

    public TransactionErrorHandler(ILoggingService logger)
    {
        _logger = logger;
    }

    public bool CanHandle(Exception exception)
    {
        return exception is TransactionException;
    }

    public ErrorHandlingResult Handle(Exception exception, Dictionary<string, object> context = null)
    {
        var transEx = exception as TransactionException;
        if (transEx == null) return null;

        _logger.LogError(transEx);

        return new ErrorHandlingResult
        {
            ShouldCancel = true,
            UserMessage = transEx.GetUserFriendlyMessage(),
            TransformedException = transEx
        };
    }
}
```

5. **Implement Sleeve Placement Error Handler (Non-Critical):**
```csharp
public class SleevePlacementErrorHandler : IErrorHandler
{
    private readonly ILoggingService _logger;

    public SleevePlacementErrorHandler(ILoggingService logger)
    {
        _logger = logger;
    }

    public bool CanHandle(Exception exception)
    {
        return exception is SleevePlacementException;
    }

    public ErrorHandlingResult Handle(Exception exception, Dictionary<string, object> context = null)
    {
        var placementEx = exception as SleevePlacementException;
        if (placementEx == null) return null;

        _logger.LogWarning(placementEx);

        return new ErrorHandlingResult
        {
            ShouldCancel = false, // Continue with next zone
            ShouldLog = true,
            UserMessage = placementEx.GetUserFriendlyMessage(),
            TransformedException = placementEx
        };
    }
}
```

### 3.3 Error Handler Chain (Chain of Responsibility Pattern)

**Files to Create:**
- `Services/ErrorHandling/ErrorHandlerChain.cs`

**Implementation Steps:**

1. **Create Error Handler Chain:**
```csharp
public class ErrorHandlerChain
{
    private readonly List<IErrorHandler> _handlers;
    private readonly ILoggingService _logger;

    public ErrorHandlerChain(ILoggingService logger)
    {
        _logger = logger;
        _handlers = new List<IErrorHandler>
        {
            new DatabaseErrorHandler(logger),
            new TransactionErrorHandler(logger),
            new RepositoryErrorHandler(logger),
            new SleevePlacementErrorHandler(logger)
        };
    }

    public ErrorHandlingResult Handle(Exception exception, Dictionary<string, object> context = null)
    {
        foreach (var handler in _handlers)
        {
            if (handler.CanHandle(exception))
            {
                return handler.Handle(exception, context);
            }
        }

        // Default handling for unhandled exceptions
        _logger.LogError("Unhandled exception", exception, context);
        return new ErrorHandlingResult
        {
            ShouldCancel = true,
            UserMessage = "An unexpected error occurred. Please contact support.",
            TransformedException = exception
        };
    }

    public void AddHandler(IErrorHandler handler)
    {
        _handlers.Add(handler);
    }
}
```

### 3.4 Result Pattern for Operations

**Files to Create:**
- `Models/OperationResult.cs`
- `Models/OperationResult<T>.cs`

**Implementation Steps:**

1. **Create Result Classes:**
```csharp
namespace JSE_RevitAddin_MEP_OPENINGS.Models
{
    public class OperationResult
    {
        public bool IsSuccess { get; private set; }
        public bool IsFailure => !IsSuccess;
        public string ErrorMessage { get; private set; }
        public Exception Exception { get; private set; }
        public Dictionary<string, object> Context { get; private set; }

        public static OperationResult Success()
        {
            return new OperationResult { IsSuccess = true };
        }

        public static OperationResult Failure(string errorMessage, Exception exception = null, Dictionary<string, object> context = null)
        {
            return new OperationResult
            {
                IsSuccess = false,
                ErrorMessage = errorMessage,
                Exception = exception,
                Context = context
            };
        }

        public static OperationResult Failure(Exception exception, Dictionary<string, object> context = null)
        {
            return new OperationResult
            {
                IsSuccess = false,
                ErrorMessage = exception.Message,
                Exception = exception,
                Context = context
            };
        }
    }

    public class OperationResult<T> : OperationResult
    {
        public T Value { get; private set; }

        public static OperationResult<T> Success(T value)
        {
            return new OperationResult<T>
            {
                IsSuccess = true,
                Value = value
            };
        }

        public new static OperationResult<T> Failure(string errorMessage, Exception exception = null, Dictionary<string, object> context = null)
        {
            return new OperationResult<T>
            {
                IsSuccess = false,
                ErrorMessage = errorMessage,
                Exception = exception,
                Context = context
            };
        }

        public new static OperationResult<T> Failure(Exception exception, Dictionary<string, object> context = null)
        {
            return new OperationResult<T>
            {
                IsSuccess = false,
                ErrorMessage = exception.Message,
                Exception = exception,
                Context = context
            };
        }
    }
}
```

### 3.5 Database Connection Error Handling (OOP)

**Files to Modify:**
- `Data/SleeveDbContext.cs`

**Implementation Steps:**

1. **Update SleeveDbContext with OOP Error Handling:**
```csharp
public class SleeveDbContext
{
    private readonly ErrorHandlerChain _errorHandler;
    private readonly ILoggingService _logger;

    public OperationResult<bool> IsConnectionValid()
    {
        try
        {
            using (var cmd = Connection.CreateCommand())
            {
                cmd.CommandText = "SELECT 1";
                cmd.ExecuteScalar();
                return OperationResult<bool>.Success(true);
            }
        }
        catch (Exception ex)
        {
            var context = new Dictionary<string, object>
            {
                { "OperationName", "Database Connection Validation" }
            };

            var result = _errorHandler.Handle(ex, context);
            return OperationResult<bool>.Failure(result.TransformedException ?? ex, context);
        }
    }

    public OperationResult<T> ExecuteWithErrorHandling<T>(
        Func<SQLiteConnection, T> operation,
        string operationName,
        Dictionary<string, object> context = null)
    {
        try
        {
            var result = operation(Connection);
            return OperationResult<T>.Success(result);
        }
        catch (Exception ex)
        {
            var errorContext = context ?? new Dictionary<string, object>();
            errorContext["OperationName"] = operationName;

            var handlingResult = _errorHandler.Handle(ex, errorContext);
            return OperationResult<T>.Failure(handlingResult.TransformedException ?? ex, errorContext);
        }
    }
}
```

**Testing:**
- Test with missing SQLite file
- Test with corrupted database
- Test with locked database

### 3.6 Filter Creation Error Handling (OOP)

**Files to Modify:**
- `Data/Repositories/FilterRepository.cs`
- `refresh refactor/refresh_service_refactored.cs`

**Implementation Steps:**

1. **Update FilterRepository with Result Pattern:**
```csharp
public class FilterRepository
{
    private readonly ErrorHandlerChain _errorHandler;
    private readonly ILoggingService _logger;

    public OperationResult<int> CreateFilter(
        string filterName, 
        string category, 
        SQLiteTransaction transaction)
    {
        try
        {
            // Existing filter creation logic
            var filterId = /* creation logic */;
            return OperationResult<int>.Success(filterId);
        }
        catch (Exception ex)
        {
            var context = new Dictionary<string, object>
            {
                { "OperationName", "CreateFilter" },
                { "FilterName", filterName },
                { "Category", category },
                { "RepositoryName", nameof(FilterRepository) },
                { "EntityType", "Filter" }
            };

            var repoException = new RepositoryException(
                nameof(FilterRepository),
                "Filter",
                "CreateFilter",
                ex,
                context);

            var handlingResult = _errorHandler.Handle(repoException, context);
            return OperationResult<int>.Failure(handlingResult.TransformedException ?? repoException, context);
        }
    }
}
```

2. **Use Result Pattern in Service:**
```csharp
var result = repo.CreateFilter(filterName, category, transaction);
if (result.IsFailure)
{
    transaction.Rollback();
    var errorResult = _errorHandler.Handle(result.Exception, result.Context);
    ShowError(errorResult.UserMessage);
    return;
}
```

**Testing:**
- Test with duplicate filter names
- Test with invalid category
- Test with transaction conflicts

### 3.7 UI State Save Error Handling (OOP)

**Files to Modify:**
- `Data/Repositories/FilterRepository.cs`
- `Services/FilterManagementService.cs`

**Implementation Steps:**

1. **Update SaveFilterUIState with Result Pattern:**
```csharp
public OperationResult SaveFilterUIState(
    string filterName, 
    string category, 
    List<string> selectedHostCategories, 
    OpeningSettings openingSettings,
    SQLiteTransaction transaction)
{
    try
    {
        // Existing save logic
        return OperationResult.Success();
    }
    catch (Exception ex)
    {
        var context = new Dictionary<string, object>
        {
            { "OperationName", "SaveFilterUIState" },
            { "FilterName", filterName },
            { "Category", category },
            { "RepositoryName", nameof(FilterRepository) }
        };

        var repoException = new RepositoryException(
            nameof(FilterRepository),
            "Filter UI State",
            "SaveFilterUIState",
            ex,
            context);

        var handlingResult = _errorHandler.Handle(repoException, context);
        return OperationResult.Failure(handlingResult.TransformedException ?? repoException, context);
    }
}
```

2. **Handle Save Failure with Error Handler:**
```csharp
var result = repo.SaveFilterUIState(filterName, category, hostCategories, settings, transaction);
if (result.IsFailure)
{
    transaction.Rollback();
    var errorResult = _errorHandler.Handle(result.Exception, result.Context);
    ShowError(errorResult.UserMessage);
    return;
}
```

**Testing:**
- Test with invalid JSON serialization
- Test with null values
- Test with very large data sets

---

## 4. Phase 2: Database Operations Enhancement (OOP)

### 4.1 ClashZone Save Error Handling (OOP)

**Files to Modify:**
- `Data/Repositories/ClashZoneRepository.cs`
- `Services/ClashZonePersistenceService.cs`

**Implementation Steps:**

1. **Create Batch Operation Result:**
```csharp
public class BatchOperationResult
{
    public int SuccessCount { get; set; }
    public int FailureCount { get; set; }
    public List<Exception> Failures { get; set; } = new List<Exception>();
    public bool IsPartialSuccess => SuccessCount > 0 && FailureCount > 0;
    public bool IsCompleteSuccess => FailureCount == 0;
    public bool IsCompleteFailure => SuccessCount == 0;
}
```

2. **Add Batch Save with OOP Error Handling:**
```csharp
public class ClashZoneRepository
{
    private readonly ErrorHandlerChain _errorHandler;
    private readonly ILoggingService _logger;

    public BatchOperationResult SaveClashZonesBatch(
        List<ClashZone> zones, 
        string filterName, 
        string category, 
        SQLiteTransaction transaction)
    {
        var result = new BatchOperationResult();

        foreach (var zone in zones)
        {
            var saveResult = SaveClashZone(zone, filterName, category, transaction);
            
            if (saveResult.IsSuccess)
            {
                result.SuccessCount++;
            }
            else
            {
                result.FailureCount++;
                result.Failures.Add(saveResult.Exception);
                
                // Use error handler for non-critical failures
                var errorResult = _errorHandler.Handle(saveResult.Exception, saveResult.Context);
                if (errorResult.ShouldLog)
                {
                    _logger.LogWarning($"Failed to save clash zone: {zone.Id}", saveResult.Exception);
                }
                // Continue with next zone (non-critical)
            }
        }

        return result;
    }

    private OperationResult SaveClashZone(
        ClashZone zone, 
        string filterName, 
        string category, 
        SQLiteTransaction transaction)
    {
        try
        {
            // Existing save logic
            return OperationResult.Success();
        }
        catch (Exception ex)
        {
            var context = new Dictionary<string, object>
            {
                { "OperationName", "SaveClashZone" },
                { "ClashZoneId", zone.Id.ToString() },
                { "FilterName", filterName },
                { "Category", category }
            };

            var repoException = new RepositoryException(
                nameof(ClashZoneRepository),
                "ClashZone",
                "SaveClashZone",
                ex,
                context);

            return OperationResult.Failure(repoException, context);
        }
    }
}
```

3. **Report Partial Success with Error Handler:**
```csharp
var batchResult = repo.SaveClashZonesBatch(zones, filterName, category, transaction);
if (batchResult.IsPartialSuccess)
{
    var errorResult = _errorHandler.Handle(
        new AggregateException(batchResult.Failures),
        new Dictionary<string, object> { { "SuccessCount", batchResult.SuccessCount } });
    
    _logger.LogWarning($"{batchResult.FailureCount} zones failed to save. {batchResult.SuccessCount} zones saved successfully.");
}
```

**Testing:**
- Test with invalid zone data
- Test with constraint violations
- Test with large batches

### 4.2 FileCombo Creation Error Handling

**Files to Modify:**
- `Data/Repositories/ClashZoneRepository.cs`

**Implementation Steps:**

1. **Add Error Handling:**
```csharp
private int GetOrCreateFileCombo(int filterId, ClashZone clashZone, SQLiteTransaction transaction)
{
    try
    {
        // Existing logic
        return comboId;
    }
    catch (SQLiteException ex) when (ex.ResultCode == SQLiteErrorCode.Constraint)
    {
        // Handle constraint violation (duplicate key)
        LogWarning($"FileCombo already exists, retrieving existing ID");
        return GetExistingFileComboId(filterId, clashZone, transaction);
    }
    catch (Exception ex)
    {
        LogError("FileCombo creation failed", ex);
        return -1;
    }
}
```

**Testing:**
- Test with duplicate file combos
- Test with invalid foreign keys
- Test with transaction conflicts

---

## 5. Phase 3: Transaction Management (OOP)

### 5.1 Transaction Manager with OOP Error Handling

**Files to Create:**
- `Data/TransactionManager.cs`
- `Services/ErrorHandling/TransactionErrorHandler.cs`

**Implementation Steps:**

1. **Create Transaction Manager with Result Pattern:**
```csharp
public class TransactionManager
{
    private readonly SleeveDbContext _context;
    private readonly ErrorHandlerChain _errorHandler;
    private readonly ILoggingService _logger;

    public OperationResult<TResult> ExecuteInTransaction<TResult>(
        Func<SQLiteTransaction, OperationResult<TResult>> operation,
        string operationName,
        Dictionary<string, object> context = null)
    {
        using (var transaction = _context.Connection.BeginTransaction())
        {
            try
            {
                var result = operation(transaction);
                
                if (result.IsFailure)
                {
                    transaction.Rollback();
                    var transException = new TransactionException(
                        operationName,
                        result.Exception,
                        wasRolledBack: true,
                        result.Context);
                    
                    var errorResult = _errorHandler.Handle(transException, result.Context);
                    return OperationResult<TResult>.Failure(errorResult.TransformedException ?? transException, result.Context);
                }

                transaction.Commit();
                _logger.LogInfo($"{operationName} completed successfully");
                return result;
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                
                var errorContext = context ?? new Dictionary<string, object>();
                errorContext["OperationName"] = operationName;
                
                var transException = new TransactionException(
                    operationName,
                    ex,
                    wasRolledBack: true,
                    errorContext);
                
                var errorResult = _errorHandler.Handle(transException, errorContext);
                return OperationResult<TResult>.Failure(errorResult.TransformedException ?? transException, errorContext);
            }
        }
    }

    public OperationResult ExecuteInTransaction(
        Func<SQLiteTransaction, OperationResult> operation,
        string operationName,
        Dictionary<string, object> context = null)
    {
        return ExecuteInTransaction(transaction => 
        {
            var result = operation(transaction);
            return OperationResult<bool>.Success(result.IsSuccess);
        }, operationName, context);
    }
}
```

2. **Use in Services with Result Pattern:**
```csharp
var transactionManager = new TransactionManager(_context, _errorHandler, _logger);

var result = transactionManager.ExecuteInTransaction(transaction =>
{
    var filterResult = repo.CreateFilter(filterName, category, transaction);
    if (filterResult.IsFailure) return filterResult;

    var uiStateResult = repo.SaveFilterUIState(filterName, category, hostCategories, settings, transaction);
    if (uiStateResult.IsFailure) return uiStateResult;

    return OperationResult.Success();
}, "Filter Creation");

if (result.IsFailure)
{
    var errorResult = _errorHandler.Handle(result.Exception, result.Context);
    ShowError(errorResult.UserMessage);
    return;
}
```

**Testing:**
- Test rollback on error
- Test commit on success
- Test nested transactions

### 5.2 Transaction Retry Logic (OOP)

**Implementation Steps:**

1. **Create Retry Strategy Interface:**
```csharp
public interface IRetryStrategy
{
    bool ShouldRetry(Exception exception, int attemptNumber, int maxRetries);
    int GetDelayMs(int attemptNumber);
}

public class DatabaseRetryStrategy : IRetryStrategy
{
    public bool ShouldRetry(Exception exception, int attemptNumber, int maxRetries)
    {
        if (attemptNumber >= maxRetries) return false;

        if (exception is SQLiteException sqlEx)
        {
            return sqlEx.ResultCode == SQLiteErrorCode.Busy ||
                   sqlEx.ResultCode == SQLiteErrorCode.Locked;
        }

        if (exception is DatabaseException dbEx && dbEx.ErrorCode.HasValue)
        {
            return dbEx.ErrorCode == SQLiteErrorCode.Busy ||
                   dbEx.ErrorCode == SQLiteErrorCode.Locked;
        }

        return false;
    }

    public int GetDelayMs(int attemptNumber)
    {
        // Exponential backoff: 100ms, 200ms, 400ms, etc.
        return 100 * (int)Math.Pow(2, attemptNumber - 1);
    }
}
```

2. **Add Retry Mechanism to Transaction Manager:**
```csharp
public class TransactionManager
{
    private readonly IRetryStrategy _retryStrategy;

    public OperationResult<TResult> ExecuteWithRetry<TResult>(
        Func<SQLiteTransaction, OperationResult<TResult>> operation,
        string operationName,
        int maxRetries = 3,
        Dictionary<string, object> context = null)
    {
        for (int attempt = 1; attempt <= maxRetries; attempt++)
        {
            var result = ExecuteInTransaction(operation, $"{operationName} (Attempt {attempt})", context);
            
            if (result.IsSuccess)
                return result;

            if (!_retryStrategy.ShouldRetry(result.Exception, attempt, maxRetries))
                return result;

            var delay = _retryStrategy.GetDelayMs(attempt);
            _logger.LogWarning($"Database busy, retrying in {delay}ms... (Attempt {attempt}/{maxRetries})");
            Thread.Sleep(delay);
        }

        return OperationResult<TResult>.Failure(
            new Exception("Transaction failed after all retries"),
            context);
    }
}
```

**Testing:**
- Test with database locks
- Test retry behavior
- Test max retry limit

---

## 6. Phase 4: Logging and Monitoring (OOP)

### 6.1 Structured Logging Interface

**Files to Create:**
- `Services/Logging/ILoggingService.cs`
- `Services/Logging/LoggingService.cs`
- `Services/Logging/LogEntry.cs`

**Implementation Steps:**

1. **Create Logging Interface:**
```csharp
public interface ILoggingService
{
    void LogError(SleevePlacementBaseException exception);
    void LogError(string message, Exception exception, Dictionary<string, object> context = null);
    void LogWarning(string message, Dictionary<string, object> context = null);
    void LogInfo(string message, Dictionary<string, object> context = null);
    void LogDebug(string message, Dictionary<string, object> context = null);
}
```

2. **Create Log Entry Model:**
```csharp
public class LogEntry
{
    public DateTime Timestamp { get; set; }
    public LogLevel Level { get; set; }
    public string Message { get; set; }
    public string OperationName { get; set; }
    public Exception Exception { get; set; }
    public Dictionary<string, object> Context { get; set; }
    public string StackTrace { get; set; }
}

public enum LogLevel
{
    Debug,
    Info,
    Warning,
    Error
}
```

3. **Implement Logging Service:**
```csharp
public class LoggingService : ILoggingService
{
    private readonly string _logDirectory;

    public LoggingService(string logDirectory = null)
    {
        _logDirectory = logDirectory ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "JSE_MEP_Openings", "Logs");
        
        Directory.CreateDirectory(_logDirectory);
    }

    public void LogError(SleevePlacementBaseException exception)
    {
        var logEntry = new LogEntry
        {
            Timestamp = exception.Timestamp,
            Level = LogLevel.Error,
            Message = exception.Message,
            OperationName = exception.OperationName,
            Exception = exception,
            Context = exception.Context,
            StackTrace = exception.StackTrace
        };

        WriteLogEntry(logEntry);
    }

    public void LogError(string message, Exception exception, Dictionary<string, object> context = null)
    {
        var logEntry = new LogEntry
        {
            Timestamp = DateTime.Now,
            Level = LogLevel.Error,
            Message = message,
            Exception = exception,
            Context = context,
            StackTrace = exception?.StackTrace
        };

        WriteLogEntry(logEntry);
    }

    public void LogWarning(string message, Dictionary<string, object> context = null)
    {
        var logEntry = new LogEntry
        {
            Timestamp = DateTime.Now,
            Level = LogLevel.Warning,
            Message = message,
            Context = context
        };

        WriteLogEntry(logEntry);
    }

    public void LogInfo(string message, Dictionary<string, object> context = null)
    {
        var logEntry = new LogEntry
        {
            Timestamp = DateTime.Now,
            Level = LogLevel.Info,
            Message = message,
            Context = context
        };

        WriteLogEntry(logEntry);
    }

    public void LogDebug(string message, Dictionary<string, object> context = null)
    {
        if (DeploymentConfiguration.DeploymentMode) return;

        var logEntry = new LogEntry
        {
            Timestamp = DateTime.Now,
            Level = LogLevel.Debug,
            Message = message,
            Context = context
        };

        WriteLogEntry(logEntry);
    }

    private void WriteLogEntry(LogEntry entry)
    {
        var logFile = Path.Combine(_logDirectory, $"sleeve_placement_{DateTime.Now:yyyy-MM-dd}.log");
        var logLine = JsonSerializer.Serialize(entry);
        
        SafeFileLogger.SafeAppendText(logFile, $"{logLine}\n");

        // Also write to debug output if not in deployment mode
        if (!DeploymentConfiguration.DeploymentMode)
        {
            switch (entry.Level)
            {
                case LogLevel.Error:
                    DebugLogger.Error($"[{entry.Level}] {entry.Message}");
                    break;
                case LogLevel.Warning:
                    DebugLogger.Warning($"[{entry.Level}] {entry.Message}");
                    break;
                case LogLevel.Info:
                    DebugLogger.Info($"[{entry.Level}] {entry.Message}");
                    break;
                case LogLevel.Debug:
                    DebugLogger.Debug($"[{entry.Level}] {entry.Message}");
                    break;
            }
        }
    }
}
```

4. **Use in Operations:**
```csharp
try
{
    // Operation
}
catch (SleevePlacementBaseException ex)
{
    _logger.LogError(ex); // Uses exception's built-in context
}
catch (Exception ex)
{
    _logger.LogError("Filter creation failed", ex, new Dictionary<string, object>
    {
        { "FilterName", filterName },
        { "Category", category }
    });
    throw;
}
```

**Testing:**
- Verify log file creation
- Verify log content
- Test log rotation

### 6.2 Performance Monitoring

**Implementation Steps:**

1. **Add Performance Tracking:**
```csharp
public class PerformanceMonitor
{
    private readonly Dictionary<string, Stopwatch> _timers = new();
    
    public void StartOperation(string operationName)
    {
        _timers[operationName] = Stopwatch.StartNew();
    }
    
    public void EndOperation(string operationName)
    {
        if (_timers.TryGetValue(operationName, out var timer))
        {
            timer.Stop();
            LogInfo($"{operationName} took {timer.ElapsedMilliseconds}ms");
            _timers.Remove(operationName);
        }
    }
}
```

2. **Use in Critical Operations:**
```csharp
_performanceMonitor.StartOperation("Filter Creation");
try
{
    // Operation
}
finally
{
    _performanceMonitor.EndOperation("Filter Creation");
}
```

**Testing:**
- Verify timing accuracy
- Test with multiple concurrent operations

---

## 7. Phase 5: User Experience Improvements

### 7.1 Progress Reporting (OOP)

**Files to Create/Modify:**
- `Services/Progress/IProgressReporter.cs`
- `Services/Progress/ProgressReporter.cs`
- `Services/Progress/ProgressEventArgs.cs`
- `Views/EmergencyMainDialog.cs`
- `Services/RefreshServiceRefactored.cs`

**Implementation Steps:**

1. **Create Progress Event Args:**
```csharp
public class ProgressEventArgs : EventArgs
{
    public int Current { get; set; }
    public int Total { get; set; }
    public string Message { get; set; }
    public double Percentage => Total > 0 ? (Current / (double)Total) * 100 : 0;
}

public class ProgressErrorEventArgs : EventArgs
{
    public string ErrorMessage { get; set; }
    public Exception Exception { get; set; }
    public bool IsCritical { get; set; }
}

public class ProgressCompleteEventArgs : EventArgs
{
    public string Summary { get; set; }
    public int SuccessCount { get; set; }
    public int FailureCount { get; set; }
}
```

2. **Create Progress Reporter Interface:**
```csharp
public interface IProgressReporter
{
    event EventHandler<ProgressEventArgs> ProgressChanged;
    event EventHandler<ProgressErrorEventArgs> ErrorOccurred;
    event EventHandler<ProgressCompleteEventArgs> Completed;

    void ReportProgress(int current, int total, string message);
    void ReportError(string message, Exception exception = null, bool isCritical = false);
    void ReportComplete(string summary, int successCount = 0, int failureCount = 0);
}
```

3. **Implement Progress Reporter:**
```csharp
public class ProgressReporter : IProgressReporter
{
    public event EventHandler<ProgressEventArgs> ProgressChanged;
    public event EventHandler<ProgressErrorEventArgs> ErrorOccurred;
    public event EventHandler<ProgressCompleteEventArgs> Completed;

    public void ReportProgress(int current, int total, string message)
    {
        var args = new ProgressEventArgs
        {
            Current = current,
            Total = total,
            Message = message
        };
        ProgressChanged?.Invoke(this, args);
    }

    public void ReportError(string message, Exception exception = null, bool isCritical = false)
    {
        var args = new ProgressErrorEventArgs
        {
            ErrorMessage = message,
            Exception = exception,
            IsCritical = isCritical
        };
        ErrorOccurred?.Invoke(this, args);
    }

    public void ReportComplete(string summary, int successCount = 0, int failureCount = 0)
    {
        var args = new ProgressCompleteEventArgs
        {
            Summary = summary,
            SuccessCount = successCount,
            FailureCount = failureCount
        };
        Completed?.Invoke(this, args);
    }
}
```

4. **Implement in Dialog (Observer Pattern):**
```csharp
public class EmergencyMainDialog : Form
{
    private readonly IProgressReporter _progressReporter;
    private ProgressBar _progressBar;
    private Label _statusLabel;

    public EmergencyMainDialog()
    {
        _progressReporter = new ProgressReporter();
        _progressReporter.ProgressChanged += OnProgressChanged;
        _progressReporter.ErrorOccurred += OnErrorOccurred;
        _progressReporter.Completed += OnCompleted;
    }

    private void OnProgressChanged(object sender, ProgressEventArgs e)
    {
        if (InvokeRequired)
        {
            Invoke(new Action(() => OnProgressChanged(sender, e)));
            return;
        }

        _progressBar.Value = (int)e.Percentage;
        _statusLabel.Text = $"{e.Message} ({e.Current}/{e.Total})";
    }

    private void OnErrorOccurred(object sender, ProgressErrorEventArgs e)
    {
        if (InvokeRequired)
        {
            Invoke(new Action(() => OnErrorOccurred(sender, e)));
            return;
        }

        if (e.IsCritical)
        {
            MessageBox.Show(e.ErrorMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        else
        {
            // Log non-critical errors
            _logger.LogWarning(e.ErrorMessage, e.Exception);
        }
    }

    private void OnCompleted(object sender, ProgressCompleteEventArgs e)
    {
        if (InvokeRequired)
        {
            Invoke(new Action(() => OnCompleted(sender, e)));
            return;
        }

        MessageBox.Show(e.Summary, "Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }
}
```

5. **Use in Operations:**
```csharp
for (int i = 0; i < zones.Count; i++)
{
    _progressReporter.ReportProgress(i + 1, zones.Count, "Placing sleeves");
    
    var result = PlaceSleeve(zones[i]);
    if (result.IsFailure)
    {
        _progressReporter.ReportError(
            $"Failed to place sleeve for zone {zones[i].Id}",
            result.Exception,
            isCritical: false);
    }
}

_progressReporter.ReportComplete(
    $"Placed {successCount} sleeves successfully. {failureCount} failed.",
    successCount,
    failureCount);
```

**Testing:**
- Test progress updates
- Test with large datasets
- Test cancellation

### 7.2 Error Messages (OOP)

**Implementation Steps:**

1. **Use Exception's Built-in User-Friendly Messages:**
```csharp
// All custom exceptions already have GetUserFriendlyMessage() method
// No need for separate helper class

catch (DatabaseException ex)
{
    var userMessage = ex.GetUserFriendlyMessage();
    MessageBox.Show(userMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
    _logger.LogError(ex);
}

catch (RepositoryException ex)
{
    var userMessage = ex.GetUserFriendlyMessage();
    MessageBox.Show(userMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
    _logger.LogError(ex);
}

catch (TransactionException ex)
{
    var userMessage = ex.GetUserFriendlyMessage();
    MessageBox.Show(userMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
    _logger.LogError(ex);
}
```

2. **Use Error Handler Chain for Automatic Message Selection:**
```csharp
catch (Exception ex)
{
    var errorResult = _errorHandler.Handle(ex, context);
    MessageBox.Show(errorResult.UserMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
    
    if (errorResult.ShouldLog)
    {
        _logger.LogError(errorResult.TransformedException ?? ex);
    }
}
```

**Testing:**
- Test all error scenarios
- Verify message clarity
- Test with non-technical users

---

## 8. Testing Strategy

### 8.1 Unit Tests

**Test Categories:**
- Database connection failures
- Transaction rollback scenarios
- Error handling in repositories
- Logging functionality

**Example Test:**
```csharp
[Test]
public void CreateFilter_RollbackOnFailure()
{
    // Arrange
    var mockContext = new Mock<SleeveDbContext>();
    var repo = new FilterRepository(mockContext.Object);
    
    // Act & Assert
    Assert.Throws<Exception>(() => repo.CreateFilter("Test", "Category", transaction));
    // Verify rollback was called
}
```

### 8.2 Integration Tests

**Test Scenarios:**
- Complete refresh flow with errors
- Database transaction integrity
- Error recovery mechanisms
- Partial success scenarios

### 8.3 User Acceptance Tests

**Test Scenarios:**
- User-friendly error messages
- Progress reporting accuracy
- System behavior under stress
- Data integrity verification

---

## 9. Deployment Checklist

### 9.1 Pre-Deployment

- [ ] All error handling code reviewed
- [ ] Unit tests passing
- [ ] Integration tests passing
- [ ] Logging configured correctly
- [ ] Error messages reviewed by non-technical users
- [ ] Performance benchmarks established
- [ ] Rollback plan prepared

### 9.2 Deployment

- [ ] Deploy to staging environment
- [ ] Run smoke tests
- [ ] Monitor error logs
- [ ] Verify transaction behavior
- [ ] Test error scenarios
- [ ] Deploy to production

### 9.3 Post-Deployment

- [ ] Monitor error rates
- [ ] Review user feedback
- [ ] Analyze performance metrics
- [ ] Update documentation
- [ ] Plan improvements

---

## 10. Phase 6: PlaceFamilyInstance Performance Optimization

### 10.1 Overview

This phase focuses on optimizing the Revit API `doc.Create.NewFamilyInstance` call and related operations to reduce sleeve placement time. The goal is **10x improvement** in placement speed.

**Current Bottlenecks:**

| Bottleneck | Impact | Current Behavior |
|------------|--------|------------------|
| **Per-Sleeve Regeneration** | ~150-300ms per sleeve | `doc.Regenerate()` called after each placement |
| **Per-Parameter Write** | ~10-50ms per param | Each `param.Set()` triggers micro-regenerations |
| **Sequential Transactions** | High BIM 360 latency | Separate transaction per category → multiple cloud syncs |
| **Family Symbol Activation** | ~50ms per symbol switch | `symbol.Activate()` called frequently |
| **Duplicate Existence Checks** | Redundant DB queries | Checking if sleeve exists before every placement |

### 10.2 Deferred Regeneration (High Impact)

**File:** `Services/NewSleevePlacerService.cs`

**Goal:** Single `doc.Regenerate()` after ALL sleeves in a category are placed, not per-sleeve.

**Implementation Steps:**

1. **Disable Regeneration During Batch:**
```csharp
using (var tx = new Transaction(doc, "Place All Sleeves"))
{
    tx.Start();
    
    // Set regeneration option to Manual
    var options = tx.GetFailureHandlingOptions();
    options.SetForcedModalHandling(false);
    tx.SetFailureHandlingOptions(options);
    
    foreach (var zone in zonesToPlace)
    {
        // Place family instance (no regen)
        var instance = doc.Create.NewFamilyInstance(...);
        
        // Set all parameters in a sub-transaction
        using (var subTx = new SubTransaction(doc))
        {
            subTx.Start();
            SetAllParameters(instance, zone);
            subTx.Commit();
        }
    }
    
    doc.Regenerate(); // Single regeneration at the end
    tx.Commit();
}
```

2. **Feature Flag:**
```csharp
public static bool UseDeferredRegeneration { get; set; } = true;
```

**Expected Improvement:** 50-70% reduction in total placement time.

### 10.3 Symbol Activation Caching (Medium Impact)

**File:** `Services/NewSleevePlacerService.cs`

**Goal:** Pre-activate ALL sleeve family symbols once at the start, not per-zone.

**Implementation Steps:**

1. **Pre-Activate All Symbols:**
```csharp
// At the start of placement batch
var symbolsNeeded = zones
    .Select(z => GetRequiredSymbol(z))
    .Distinct()
    .ToList();

foreach (var symbol in symbolsNeeded)
{
    if (!symbol.IsActive)
    {
        symbol.Activate();
        doc.Regenerate(); // Required after activation
    }
}

// Cache in dictionary for O(1) lookup during placement
var symbolCache = symbolsNeeded.ToDictionary(
    s => GetSymbolKey(s),
    s => s);
```

2. **Group Zones by Symbol:**
```csharp
// Sort zones by symbol to minimize any remaining switching overhead
var zonesBySymbol = zones
    .GroupBy(z => GetSymbolKey(z))
    .OrderBy(g => g.Key)
    .SelectMany(g => g)
    .ToList();
```

**Expected Improvement:** 10-15% reduction in total placement time.

### 10.4 Transaction Consolidation (BIM 360 Critical)

**File:** `Services/OpeningCommandOrchestrator.cs`

**Goal:** Single transaction per discipline (Ducts + Pipes + Cable Trays), not per category.

**Implementation Steps:**

1. **Unified Placement Transaction:**
```csharp
// Instead of:
// foreach (filter in filters) { ExecutePerFilter(filter); }

// Use:
if (OptimizationFlags.UseUnifiedAllCategoryPlacement)
{
    using (var tx = new Transaction(doc, "Place All Discipline Sleeves"))
    {
        tx.Start();
        
        foreach (var filter in filters)
        {
            var zones = LoadZonesForFilter(filter);
            PlaceSleevesWithoutTransaction(zones);
        }
        
        doc.Regenerate(); // Single regen for entire discipline
        tx.Commit();
    }
}
```

2. **Feature Flag (Already Exists):**
```csharp
public static bool UseUnifiedAllCategoryPlacement { get; set; } = true;
```

**Expected Improvement:** 5-10x improvement for BIM 360 projects (reduced cloud sync calls).

### 10.5 Parallel Pre-Calculation (CPU-Bound)

**Files:** 
- `Services/Placement/IndividualSleevePreCalculationService.cs`
- `Services/Parallel/ParallelSleevePlacementPlanner.cs`

**Goal:** Pre-calculate dimensions, clearance, and rotation off the main Revit thread.

**Implementation Steps:**

1. **Parallel Calculation:**
```csharp
// These calculations do NOT use Revit API, so they can run in parallel
Parallel.ForEach(zones, zone =>
{
    // Pure math - no Revit API calls
    zone.PreCalculatedWidth = CalculateWidth(zone);
    zone.PreCalculatedHeight = CalculateHeight(zone);
    zone.PreCalculatedRotationRad = CalculateRotation(zone);
    zone.PreCalculatedOffset = CalculateOffset(zone);
});
```

2. **Use Pre-Calculated Values During Placement:**
```csharp
// Inside placement loop - just read values, no calculation
var width = zone.PreCalculatedWidth;
var height = zone.PreCalculatedHeight;
var rotation = zone.PreCalculatedRotationRad;

var instance = doc.Create.NewFamilyInstance(...);
// Set parameters using pre-calculated values
```

3. **Feature Flag (Already Exists):**
```csharp
public static bool UseSOLIDRefactoredIndividualPreCalculation { get; set; } = true;
```

**Expected Improvement:** 20-30% reduction in main thread time.

### 10.6 Existence Check Elimination (Quick Win)

**File:** `Services/NewSleevePlacerService.cs`

**Goal:** Trust database flags instead of querying Revit for existing sleeves.

**Implementation Steps:**

1. **Trust Database Flags:**
```csharp
// Instead of:
// if (doc.GetElement(new ElementId(zone.SleeveInstanceId)) != null) continue;

// Use:
if (zone.IsResolvedFlag || zone.SleeveInstanceId > 0)
{
    continue; // Already placed - trust the database
}
```

2. **Optional: Batch Existence Check:**
```csharp
// If paranoid, do ONE collector query at the start
var existingSleeveIds = new FilteredElementCollector(doc)
    .OfClass(typeof(FamilyInstance))
    .OfCategory(BuiltInCategory.OST_GenericModel)
    .WhereElementIsNotElementType()
    .Select(e => e.Id.IntegerValue)
    .ToHashSet();

// Use HashSet for O(1) lookup instead of per-zone API call
if (existingSleeveIds.Contains(zone.SleeveInstanceId)) continue;
```

**Expected Improvement:** 5-10% reduction in placement overhead.

### 10.7 Feature Flag Summary

| Flag | Purpose | Production | Debug |
|------|---------|------------|-------|
| `UseUnifiedAllCategoryPlacement` | Single transaction per discipline | `true` | `false` |
| `UseSOLIDRefactoredIndividualPreCalculation` | Parallel pre-calculation | `true` | `false` |
| `UseSOLIDRefactoredClusterPreCalculation` | Parallel cluster pre-calc | `true` | `false` |
| `UseDeferredRegeneration` | Single regen per batch | `true` | `false` |
| `SuppressPlacementPrompts` | Skip per-category dialogs | `true` | `false` |

### 10.8 Verification Plan

**Metrics to Track:**
- Total Placement Time (all categories)
- Per-Sleeve Time (average)
- Transaction Count (should be 1 for unified)
- Regeneration Count (should be 1 per category)

**Log Files:**
- `placement_perf.log` - Per-category timing
- `orchestrator_debug.log` - Transaction boundaries
- `planning_debug.log` - Parallel pre-calculation timings

**Success Criteria:**
- ✅ **10x improvement** in placement speed (target: <100ms per sleeve)
- ✅ **Single cloud sync** per discipline on BIM 360
- ✅ **No regressions** in placement accuracy

### 10.9 Bulk Placement Data Model (Database-Driven)

**Goal:** Query ALL placement data from SQLite database in ONE query, build a `List<SleevePlacementRequest>`, and execute bulk placement using pre-computed values.

#### 10.9.1 SleevePlacementRequest DTO

**New File:** `Models/SleevePlacementRequest.cs`

```csharp
/// <summary>
/// Immutable DTO containing all data needed to place a single sleeve.
/// Populated from ClashZone database fields during pre-loading phase.
/// </summary>
public class SleevePlacementRequest
{
    // ===== IDENTIFICATION =====
    public Guid ClashZoneId { get; init; }
    public int ClashZoneDbId { get; init; }  // SQLite auto-increment ID for fast updates
    
    // ===== FAMILY SELECTION =====
    /// <summary>Which family to use (e.g., "RectangularOpeningOnWall", "RoundOpeningOnFloor")</summary>
    public string FamilyName { get; init; }
    
    /// <summary>Host type: "Wall", "Floor", "Structural Framing"</summary>
    public string HostType { get; init; }
    
    /// <summary>Host orientation: "X" or "Y" for walls, empty for floors</summary>
    public string HostOrientation { get; init; }
    
    /// <summary>Duct shape: "Round" or "Rectangular"</summary>
    public string DuctShape { get; init; }
    
    /// <summary>MEP category: "Ducts", "Pipes", "Cable Trays", "Duct Accessories"</summary>
    public string MepCategory { get; init; }
    
    // ===== PLACEMENT POINT (World Coordinates) =====
    public double PlacementX { get; init; }
    public double PlacementY { get; init; }
    public double PlacementZ { get; init; }
    
    // ===== SLEEVE DIMENSIONS (Revit Internal Units - Feet) =====
    public double Width { get; init; }   // For rectangular sleeves
    public double Height { get; init; }  // For rectangular sleeves
    public double Diameter { get; init; } // For round sleeves
    public double Depth { get; init; }   // Wall/floor thickness
    
    // ===== ROTATION (Radians - For Floor Sleeves) =====
    public double RotationAngleRad { get; init; }
    
    // ===== PRE-CALCULATED HELPERS =====
    public double RotationCos { get; init; }
    public double RotationSin { get; init; }
    
    // ===== LEVEL REFERENCE (For Revit Placement) =====
    public string LevelName { get; init; }
    public double LevelElevation { get; init; }
    
    // ===== PARAMETER SNAPSHOT (For Post-Placement Transfer) =====
    public string MepParameterValuesJson { get; init; }
    public string HostParameterValuesJson { get; init; }
}
```

#### 10.9.2 Database Fields Mapping

| SleevePlacementRequest | ClashZone Property (Database Column) |
|------------------------|--------------------------------------|
| `FamilyName` | `SleeveFamilyName` |
| `HostType` | `StructuralElementType` |
| `HostOrientation` | `HostOrientation` |
| `DuctShape` | `DuctShape` |
| `MepCategory` | `MepElementCategory` |
| `PlacementX/Y/Z` | `SleevePlacementPointX/Y/Z` |
| `Width` | `SleeveWidth` |
| `Height` | `SleeveHeight` |
| `Diameter` | `SleeveDiameter` |
| `Depth` | `StructuralElementThickness` |
| `RotationAngleRad` | `MepElementRotationAngle` |
| `RotationCos/Sin` | `MepRotationCos`/`MepRotationSin` |
| `LevelName` | `MepElementLevelName` |
| `LevelElevation` | `MepElementLevelElevation` |
| `MepParameterValuesJson` | `MepParameterValuesJson` |
| `HostParameterValuesJson` | `HostParameterValuesJson` |

#### 10.9.3 Unified SQL Query for ALL Categories

**Repository Method:** `GetAllPlacementRequestsForBulk()`

```sql
-- ✅ UNIFIED QUERY: Get ALL zones ready for placement (all categories at once)
-- ReadyForPlacementFlag is ONLY set for unresolved zones (see SetReadyForPlacementForUnresolvedZonesInSectionBox)
-- So no need to check IsResolvedFlag, IsClusterResolvedFlag, or IsCombinedResolved again
SELECT 
    ClashZoneGuid,
    ClashZoneId,
    SleeveFamilyName,
    StructuralElementType,
    HostOrientation,
    DuctShape,
    MepElementCategory,
    SleevePlacementPointX,
    SleevePlacementPointY,
    SleevePlacementPointZ,
    SleeveWidth,
    SleeveHeight,
    SleeveDiameter,
    StructuralElementThickness,
    MepElementRotationAngle,
    MepRotationCos,
    MepRotationSin,
    MepElementLevelName,
    MepElementLevelElevation
FROM ClashZones
WHERE IsCurrentClashFlag = 1           -- Active in current session (section box + filters)
  AND ReadyForPlacementFlag = 1        -- Unresolved + ready to place
ORDER BY MepElementCategory, SleeveFamilyName  -- Group by category for symbol caching
```

**Why only 2 flags?**
- `SetReadyForPlacementForUnresolvedZonesInSectionBox` already applies: `IsResolvedFlag=0 AND IsClusterResolvedFlag=0 AND IsCombinedResolved=0`
- If `ReadyForPlacementFlag = 1`, the zone is **guaranteed** unresolved
- No redundant checks needed!

#### 10.9.4 Complete Bulk Placement Flow

```
USER CLICKS "PLACE ALL SLEEVES"
    │
    ├─ STEP 1: Query DB (raw MEP data)
    │     SELECT * FROM ClashZones
    │     WHERE IsCurrentClashFlag=1 AND ReadyForPlacementFlag=1
    │
    ├─ STEP 2: Calculate final sizes using CURRENT UI settings
    │     For each zone:
    │       SleeveWidth = MepWidth + (2 × ClearanceFromUI) + InsulationThickness
    │       SleeveHeight = MepHeight + (2 × ClearanceFromUI) + InsulationThickness
    │       SleeveFamilyName = DetermineFamilyName(HostType, DuctShape)
    │       RotationAngle = MepElementRotationAngle (for floors)
    │
    ├─ STEP 3: SAVE calculated sizes to DB ✅
    │     UPDATE ClashZones SET 
    │       SleeveWidth = @width,
    │       SleeveHeight = @height,
    │       SleeveDiameter = @diameter,
    │       SleeveFamilyName = @familyName
    │     WHERE ClashZoneId IN (@ids)
    │
    ├─ STEP 4: Build List<FamilyInstanceCreationData> (from DB data)
    │     Each entry has: Point, Symbol (from FamilyName), StructuralType
    │
    ├─ STEP 5: Single Transaction - Bulk Placement
    │     createdIds = doc.Create.NewFamilyInstances2(creationDataList)
    │     For each created instance:
    │       Set Width/Height/Diameter parameters
    │       Apply Rotation (for floors)
    │     doc.Regenerate() once at the end
    │
    └─ STEP 6: Batch update DB with placement results
          UPDATE ClashZones SET 
            SleeveInstanceId = @instanceId,
            IsResolvedFlag = 1,
            ReadyForPlacementFlag = 0
          WHERE ClashZoneId = @id
```

#### 10.9.5 Implementation Code

```csharp
public void ExecuteBulkPlacement(Document doc)
{
    // STEP 1: Query DB
    List<ClashZone> zones;
    using (var dbContext = new SleeveDbContext(doc))
    {
        var repo = new ClashZoneRepository(dbContext);
        zones = repo.GetZonesReadyForPlacement(); // IsCurrentClash=1 AND ReadyForPlacement=1
    }
    
    if (zones.Count == 0) return;
    
    // STEP 2: Calculate sizes and family names using CURRENT UI settings
    // Uses existing logic which handles:
    //   - Clearance per category (from UI)
    //   - Round → Rectangular threshold (from UI, e.g., >200mm)
    //   - Host type (Wall/Framing → *OnWall, Floor → *OnSlab)
    //   - MEP element shape (circular vs rectangular)
    var uiSettings = OpeningSettingsHelper.GetClearanceSettings();
    var shapeThreshold = OpeningSettingsHelper.GetRoundToRectangularThreshold();
    
    // =====================================================
    // ✅ SAFE PARALLEL PROCESSING STRATEGY
    // =====================================================
    // PRINCIPLE: Sequential is default, parallel is opt-in for large batches
    // THRESHOLD: Only parallelize when zones.Count > 500 (big projects, 10+ floors)
    // SAFETY: Each thread processes INDEPENDENT zones - no shared state mutation
    
    bool useParallel = OptimizationFlags.UseSOLIDRefactoredIndividualPreCalculation 
                       && zones.Count > 500;  // High threshold = safer
    
    if (useParallel)
    {
        // ✅ SAFE PATTERN: Use ConcurrentBag for thread-safe result collection
        // Each zone's properties are ONLY modified by ONE thread (index-based partitioning)
        // No two threads ever touch the same zone
        Parallel.ForEach(
            Partitioner.Create(0, zones.Count),
            new ParallelOptions { MaxDegreeOfParallelism = Math.Max(2, Environment.ProcessorCount / 2) },
            range =>
            {
                // This thread ONLY processes zones[range.Item1] to zones[range.Item2 - 1]
                // Other threads process OTHER index ranges
                for (int i = range.Item1; i < range.Item2; i++)
                {
                    var zone = zones[i];  // Thread owns this zone exclusively
                    CalculateZoneSizeAndFamily(zone, uiSettings, shapeThreshold);
                }
            });
        
        SafeFileLogger.SafeAppendText("bulk_placement.log", 
            $"[{DateTime.Now:HH:mm:ss}] ✅ PARALLEL: {zones.Count} zones processed in parallel\n");
    }
    else
    {
        // ✅ SEQUENTIAL: Safe, predictable, easy to debug
        foreach (var zone in zones)
        {
            CalculateZoneSizeAndFamily(zone, uiSettings, shapeThreshold);
        }
        
        SafeFileLogger.SafeAppendText("bulk_placement.log", 
            $"[{DateTime.Now:HH:mm:ss}] ✅ SEQUENTIAL: {zones.Count} zones processed sequentially\n");
    }

// =====================================================
// HELPER METHOD - Thread-safe by design
// =====================================================
// SAFETY: This method ONLY modifies properties of the zone passed to it
// No shared state, no static variables, no external collections modified
void CalculateZoneSizeAndFamily(ClashZone zone, ClearanceSettings settings, double threshold)
{
    // All values calculated from zone's own properties + read-only settings
    double clearance = GetClearance(zone.MepElementCategory, zone.IsInsulated, settings);
    zone.SleeveWidth = zone.MepElementWidth + (2 * clearance) + (zone.IsInsulated ? zone.InsulationThickness * 2 : 0);
    zone.SleeveHeight = zone.MepElementHeight + (2 * clearance) + (zone.IsInsulated ? zone.InsulationThickness * 2 : 0);
    
    var (_, _, diameter, isCircular) = CalculateSleeveDimensions(zone);
    zone.SleeveDiameter = diameter;
    zone.SleeveFamilyName = GetSleeveFamilyName(zone, isCircular);
}
    
    // STEP 3: SAVE calculated sizes to DB
    using (var dbContext = new SleeveDbContext(doc))
    {
        var repo = new ClashZoneRepository(dbContext);
        repo.BatchUpdateSleeveSizes(zones); // Saves Width, Height, Diameter, FamilyName
    }
    
    // STEP 4: Pre-activate symbols and build creation data
    var symbolCache = PreActivateSymbols(doc, zones);
    var creationDataList = zones.Select(z => new FamilyInstanceCreationData(
        new XYZ(z.SleevePlacementPointX, z.SleevePlacementPointY, z.SleevePlacementPointZ),
        symbolCache[z.SleeveFamilyName],
        StructuralType.NonStructural
    )).ToList();
    
    // STEP 5: Single Transaction - Bulk Placement
    using (var tx = new Transaction(doc, "Bulk Sleeve Placement"))
    {
        tx.Start();
        
        var createdIds = doc.Create.NewFamilyInstances2(creationDataList);
        
        // Set parameters on all created instances
        var idList = createdIds.ToList();
        for (int i = 0; i < idList.Count; i++)
        {
            var sleeve = doc.GetElement(idList[i]) as FamilyInstance;
            var zone = zones[i];
            
            sleeve.LookupParameter("Width")?.Set(zone.SleeveWidth);
            sleeve.LookupParameter("Height")?.Set(zone.SleeveHeight);
            if (zone.SleeveDiameter > 0)
                sleeve.LookupParameter("Diameter")?.Set(zone.SleeveDiameter);
            
            // Apply rotation for floors
            if (zone.StructuralElementType == "Floor" && Math.Abs(zone.MepElementRotationAngle) > 0.001)
            {
                var point = (sleeve.Location as LocationPoint).Point;
                var axis = Line.CreateBound(point, point.Add(XYZ.BasisZ));
                ElementTransformUtils.RotateElement(doc, sleeve.Id, axis, zone.MepElementRotationAngle);
            }
            
            // Track for DB update
            zone.SleeveInstanceId = sleeve.Id.IntegerValue;
        }
        
        doc.Regenerate();
        tx.Commit();
    }
    
    // STEP 6: Batch update DB with placement results
    using (var dbContext = new SleeveDbContext(doc))
    {
        var repo = new ClashZoneRepository(dbContext);
        repo.BatchUpdateAfterPlacement(zones); // Sets SleeveInstanceId, IsResolved=1, ReadyForPlacement=0
    }
}
```

#### 10.9.6 Key Benefits of This Approach

| Aspect | Before | After |
|--------|--------|-------|
| **DB Queries** | 1 per zone | 1 for ALL zones |
| **Symbol Activation** | Per-zone | Once upfront |
| **Revit Transactions** | Per-category | 1 for ALL categories |
| **Regeneration** | Per-sleeve | 1 at the end |
| **Dimension Calculation** | During placement | Pre-loaded from DB |
| **Level Lookup** | Per-zone | Cached once |

#### 10.9.7 Family Name Determination Logic (DetermineOpeningType)

**Critical:** The `isCircular` boolean that determines family name involves complex rules, not just diameter check.

**Method:** `DetermineOpeningType(zone, rawDiameter, finalDiameter)` (NewSleevePlacerService.cs:1452-1554)

```
┌─────────────────────────────────────────────────────────────────┐
│                    DetermineOpeningType()                        │
└─────────────────────────────────────────────────────────────────┘
                              │
        ┌─────────────────────┼─────────────────────┐
        ▼                     ▼                     ▼
┌───────────────┐    ┌───────────────┐    ┌───────────────┐
│    PIPES      │    │    DUCTS      │    │    OTHER      │
│               │    │               │    │ (Cable Trays) │
│ 1. MepSize obj│    │ 1. DuctShape  │    │               │
│ 2. UI Pref    │    │ 2. UI Pref    │    │ → FALSE       │
│ 3. GetResolved│    │ 3. Round      │    │   (always)    │
│    OpeningType│    │    Ducts pref │    │               │
│ 4. Threshold! │    │               │    │               │
└───────────────┘    └───────────────┘    └───────────────┘
```

**PIPES Category Rules:**
1. Create `MepElementSize` object from ClashZone
2. Get UI preference: `_conditions?.OpeningTypePreferences?.Pipes` (from CONDITIONS XML)
3. Call `PipePlacementStrategy.GetResolvedOpeningType()`:
   - Checks `RoundOpeningsBecomeRectangularIfDiameterGreaterThan` threshold (default 200mm from UI)
   - Structural framing rule: pipes on framing → always circular
4. Fallback: `rawDiameter > 0`

**DUCTS Category Rules:**
1. Check `zone.DuctShape` == "Round" or "Circular"
2. If round: Use `_conditions?.OpeningTypePreferences?.RoundDucts` (UI setting)
3. If rectangular: Always `isCircular = false`

**OTHER Categories (Cable Trays, etc.):** Always `isCircular = false`

**Performance Impact:**
- Object creation (`MepElementSize`)
- Strategy pattern method call
- Multiple string comparisons with logging
- **Estimated: ~0.1-0.5ms per zone** (not trivial)

**For Bulk Placement:** Must call `DetermineOpeningType()` for each zone to correctly determine family name.

### 10.11 SOLID Implementation with Safe Transactions & Crash Recovery

#### 10.11.1 SOLID Architecture for Bulk Placement

```
┌─────────────────────────────────────────────────────────────────┐
│                    IBulkPlacementService                        │
│  (Single Responsibility: Orchestrate bulk placement flow)       │
└─────────────────────────────────────────────────────────────────┘
                              │
        ┌─────────────────────┼─────────────────────┐
        ▼                     ▼                     ▼
┌───────────────┐    ┌───────────────┐    ┌───────────────┐
│ISizeCalculator│    │ISymbolManager │    │IPlacementExec │
│(Pure math,    │    │(Load, activate│    │(NewFamilyInst │
│ no Revit API) │    │ cache symbols)│    │ ances2 call)  │
└───────────────┘    └───────────────┘    └───────────────┘
        │                     │                     │
        ▼                     ▼                     ▼
┌───────────────┐    ┌───────────────┐    ┌───────────────┐
│IDbPersistence │    │ITransactionMgr│    │IRecoveryMgr   │
│(Batch DB ops) │    │(Safe commits) │    │(Rollback/retry│
└───────────────┘    └───────────────┘    └───────────────┘
```

#### 10.11.2 Safe Transaction Management

```csharp
public class SafeBulkPlacementService : IBulkPlacementService
{
    public BulkPlacementResult ExecuteSafeBulkPlacement(Document doc, List<ClashZone> zones)
    {
        var result = new BulkPlacementResult();
        var placedZones = new List<ClashZone>();  // Track for rollback
        
        try
        {
            // ═══════════════════════════════════════════════════════════
            // PHASE 1: PRE-FLIGHT VALIDATION (Before any changes)
            // ═══════════════════════════════════════════════════════════
            if (!ValidateDocument(doc))
            {
                result.Error = "Document not modifiable";
                return result;
            }
            
            if (zones == null || zones.Count == 0)
            {
                result.Error = "No zones to place";
                return result;
            }
            
            // ═══════════════════════════════════════════════════════════
            // PHASE 2: SIZE CALCULATION (No Revit API, safe to fail)
            // ═══════════════════════════════════════════════════════════
            try
            {
                CalculateSizesAndFamilyNames(zones);
            }
            catch (Exception calcEx)
            {
                result.Error = $"Size calculation failed: {calcEx.Message}";
                return result;  // No cleanup needed - nothing modified yet
            }
            
            // ═══════════════════════════════════════════════════════════
            // PHASE 3: SAVE TO DB (Atomic, can rollback)
            // ═══════════════════════════════════════════════════════════
            try
            {
                BatchSaveToDatabase(zones);
            }
            catch (Exception dbEx)
            {
                result.Error = $"Database save failed: {dbEx.Message}";
                return result;  // DB transaction rolled back automatically
            }
            
            // ═══════════════════════════════════════════════════════════
            // PHASE 4: SYMBOL ACTIVATION (Separate transaction)
            // ═══════════════════════════════════════════════════════════
            Dictionary<string, FamilySymbol> symbolCache;
            try
            {
                symbolCache = PreActivateSymbolsSafe(doc, zones);
            }
            catch (Exception symEx)
            {
                result.Error = $"Symbol activation failed: {symEx.Message}";
                // Rollback: Clear size data from DB
                ClearCalculatedSizes(zones);
                return result;
            }
            
            // ═══════════════════════════════════════════════════════════
            // PHASE 5: BULK PLACEMENT (Main transaction with rollback)
            // ═══════════════════════════════════════════════════════════
            using (var transactionGroup = new TransactionGroup(doc, "Bulk Sleeve Placement"))
            {
                transactionGroup.Start();
                
                try
                {
                    using (var tx = new Transaction(doc, "Place All Sleeves"))
                    {
                        tx.Start();
                        
                        // ✅ CRASH-SAFE: Validate document before placement
                        if (!doc.IsModifiable)
                        {
                            tx.RollBack();
                            throw new InvalidOperationException("Document became read-only");
                        }
                        
                        // Build creation data list
                        var creationDataList = BuildCreationDataList(zones, symbolCache);
                        
                        // ✅ THE KEY API CALL
                        var createdIds = doc.Create.NewFamilyInstances2(creationDataList);
                        
                        // ✅ CRASH-SAFE: Validate creation succeeded
                        if (createdIds.Count != zones.Count)
                        {
                            tx.RollBack();
                            throw new Exception($"Expected {zones.Count} instances, got {createdIds.Count}");
                        }
                        
                        // Set parameters and track results
                        var idList = createdIds.ToList();
                        for (int i = 0; i < idList.Count; i++)
                        {
                            var sleeve = doc.GetElement(idList[i]) as FamilyInstance;
                            
                            // ✅ CRASH-SAFE: Validate each element
                            if (sleeve == null || !sleeve.IsValidObject)
                            {
                                SafeFileLogger.SafeAppendText("bulk_errors.log",
                                    $"[{DateTime.Now:HH:mm:ss}] ⚠️ Sleeve at index {i} is null/invalid\n");
                                continue;
                            }
                            
                            SetSleeveParameters(sleeve, zones[i]);
                            zones[i].SleeveInstanceId = sleeve.Id.IntegerValue;
                            placedZones.Add(zones[i]);
                        }
                        
                        doc.Regenerate();
                        tx.Commit();
                    }
                    
                    // ✅ SUCCESS: Assimilate all transactions
                    transactionGroup.Assimilate();
                    result.Success = true;
                    result.PlacedCount = placedZones.Count;
                }
                catch (Exception txEx)
                {
                    // ✅ CRASH-SAFE: Rollback entire group
                    transactionGroup.RollBack();
                    result.Error = $"Placement failed: {txEx.Message}";
                    
                    // Clear SleeveInstanceIds that were set
                    foreach (var zone in placedZones)
                    {
                        zone.SleeveInstanceId = 0;
                    }
                    
                    return result;
                }
            }
            
            // ═══════════════════════════════════════════════════════════
            // PHASE 6: UPDATE DATABASE (After successful Revit commit)
            // ═══════════════════════════════════════════════════════════
            try
            {
                BatchUpdateDatabaseAfterPlacement(placedZones);
            }
            catch (Exception dbUpdateEx)
            {
                // ⚠️ Sleeves placed in Revit but DB not updated
                // Log for manual recovery
                SafeFileLogger.SafeAppendText("bulk_recovery.log",
                    $"[{DateTime.Now:HH:mm:ss}] ⚠️ DB update failed after placement: {dbUpdateEx.Message}\n" +
                    $"Placed {placedZones.Count} sleeves - IDs: {string.Join(",", placedZones.Select(z => z.SleeveInstanceId))}\n");
                result.Warning = "Sleeves placed but DB update failed - check recovery log";
            }
        }
        catch (Exception ex)
        {
            result.Error = $"Unexpected error: {ex.Message}";
            SafeFileLogger.SafeAppendText("bulk_errors.log",
                $"[{DateTime.Now:HH:mm:ss}] ❌ CRITICAL: {ex.Message}\n{ex.StackTrace}\n");
        }
        
        return result;
    }
}
```

#### 10.11.3 Crash Recovery Features

| Scenario | Detection | Recovery Action |
|----------|-----------|-----------------|
| **Calculation fails** | Exception in Phase 2 | Return immediately, nothing modified |
| **DB save fails** | Exception in Phase 3 | DB transaction rolls back automatically |
| **Symbol activation fails** | Exception in Phase 4 | Clear calculated sizes from zones |
| **Placement fails mid-batch** | Exception in Phase 5 | `TransactionGroup.RollBack()` undoes all Revit changes |
| **Revit crashes** | On next startup | DB has sizes saved, can retry placement |
| **DB update fails after placement** | Exception in Phase 6 | Log sleeve IDs for manual recovery |

#### 10.11.4 Result Pattern

```csharp
public class BulkPlacementResult
{
    public bool Success { get; set; }
    public int PlacedCount { get; set; }
    public int SkippedCount { get; set; }
    public string? Error { get; set; }
    public string? Warning { get; set; }
    public List<(Guid ZoneId, string Reason)> Failures { get; set; } = new();
}
```

### 10.10 Official Revit API: `NewFamilyInstances2` (Batch Creation)

> [!IMPORTANT]
> **This is the industry-standard approach for bulk family placement.** Revit provides a dedicated API specifically for batch creation that is significantly faster than looped `NewFamilyInstance` calls.

#### 10.10.1 The Official API

**Method:** `Document.Create.NewFamilyInstances2(IList<FamilyInstanceCreationData>)`

**Returns:** `ICollection<ElementId>` - IDs of all created instances

**Key Benefits:**
- **Single API call** for ALL instances (not N calls)
- **Optimized internal batching** - Revit handles memory and regeneration internally
- **Reduced overhead** - Avoids per-element commit overhead
- **Significant speedup** - Developers report hours → minutes improvement

#### 10.10.2 FamilyInstanceCreationData for Non-Hosted Generic Models

Your sleeves are **Generic Model** category, **non-hosted**. You have TWO patterns in your codebase:

**Pattern 1: Individual Sleeves (WITHOUT Level) - Absolute Z Coordinate**
```csharp
// From NewSleevePlacerService.cs line 1739:
FamilyInstance instance = _doc.Create.NewFamilyInstance(point, symbol, StructuralType.NonStructural);

// Batch equivalent:
var creationData = new FamilyInstanceCreationData(
    point,                             // ABSOLUTE world coordinates (Z is exact elevation)
    symbol,                            // Pre-activated FamilySymbol
    StructuralType.NonStructural       // Always NonStructural for Generic Models
);
// Then set Level parameter AFTER creation via INSTANCE_REFERENCE_LEVEL_PARAM
```

**Pattern 2: Cluster Sleeves (WITH Level) - Level-Relative Z Coordinate**
```csharp
// From ClusterPlacementService.cs line 385:
inst = doc.Create.NewFamilyInstance(placementPoint, familySymbol, refLevel, StructuralType.NonStructural);

// Batch equivalent:
var creationData = new FamilyInstanceCreationData(
    point,                             // Level-relative coordinates
    symbol,                            // Pre-activated FamilySymbol
    level,                             // Reference level (Z will be relative to this)
    StructuralType.NonStructural
);
```

> [!IMPORTANT]
> **Choose based on your placement point source:**
> - If `SleevePlacementPointZ` is **ABSOLUTE** (already in world coordinates) → Use Pattern 1 (no level)
> - If `SleevePlacementPointZ` is **level-relative** (stored as offset from level) → Use Pattern 2 (with level)
> 
> Your database stores **ABSOLUTE** coordinates, so **Pattern 1 is correct for bulk placement**.

> [!NOTE]
> **Cluster placement (Pattern 2) works correctly** as a separate code path. The midpoint calculation in `ComputeClusterMidpoint()` produces coordinates that work with the 4-param overload. Don't change it if it's working.

#### 10.10.3 Complete Batch Placement Implementation

```csharp
public ICollection<ElementId> ExecuteBulkPlacementWithRevitAPI(
    Document doc, 
    List<SleevePlacementRequest> requests)
{
    // PHASE 1: Pre-activate all required symbols
    var symbolCache = new Dictionary<string, FamilySymbol>();
    using (var tx = new Transaction(doc, "Activate Symbols"))
    {
        tx.Start();
        foreach (var familyName in requests.Select(r => r.FamilyName).Distinct())
        {
            var symbol = GetSymbolByName(doc, familyName);
            if (!symbol.IsActive)
            {
                symbol.Activate();
            }
            symbolCache[familyName] = symbol;
        }
        tx.Commit();
    }

    // PHASE 2: Build FamilyInstanceCreationData list (NO Revit API calls here)
    var creationDataList = new List<FamilyInstanceCreationData>();
    
    foreach (var request in requests)
    {
        var point = new XYZ(request.PlacementX, request.PlacementY, request.PlacementZ);
        var symbol = symbolCache[request.FamilyName];
        
        // ✅ CRITICAL: Use 3-parameter constructor (point, symbol, structuralType)
        // DO NOT pass Level - you want ABSOLUTE Z coordinate placement
        var creationData = new FamilyInstanceCreationData(
            point,
            symbol,
            StructuralType.NonStructural  // Always NonStructural for Generic Model sleeves
        );
        
        creationDataList.Add(creationData);
    }

    // PHASE 3: Single Transaction, Single API Call
    ICollection<ElementId> createdIds;
    using (var tx = new Transaction(doc, "Bulk Sleeve Placement"))
    {
        tx.Start();
        
        // ✅ THE KEY API CALL - Creates ALL instances at once
        createdIds = doc.Create.NewFamilyInstances2(creationDataList);
        
        // ✅ Set parameters AFTER all instances are created
        var createdIdsList = createdIds.ToList();
        for (int i = 0; i < createdIdsList.Count; i++)
        {
            var sleeve = doc.GetElement(createdIdsList[i]) as FamilyInstance;
            var request = requests[i];
            
            // Set Width/Height/Diameter/Depth from pre-loaded DB values
            SetSleeveParameters(sleeve, request);
            
            // Apply rotation for floor sleeves
            if (request.HostType == "Floor" && Math.Abs(request.RotationAngleRad) > 0.001)
            {
                var axis = Line.CreateBound(
                    (sleeve.Location as LocationPoint).Point, 
                    (sleeve.Location as LocationPoint).Point.Add(XYZ.BasisZ));
                ElementTransformUtils.RotateElement(doc, sleeve.Id, axis, request.RotationAngleRad);
            }
        }
        
        tx.Commit();
    }

    return createdIds;
}

private void SetSleeveParameters(FamilyInstance sleeve, SleevePlacementRequest request)
{
    // Width/Height for rectangular sleeves
    if (request.Width > 0)
        sleeve.LookupParameter("Width")?.Set(request.Width);
    if (request.Height > 0)
        sleeve.LookupParameter("Height")?.Set(request.Height);
    
    // Diameter for round sleeves
    if (request.Diameter > 0)
        sleeve.LookupParameter("Diameter")?.Set(request.Diameter);
    
    // Depth (wall/floor thickness)
    if (request.Depth > 0)
        sleeve.LookupParameter("Depth")?.Set(request.Depth);
}
```

#### 10.10.4 Performance Comparison

| Approach | 100 Sleeves | 500 Sleeves | 1000 Sleeves |
|----------|-------------|-------------|--------------|
| **Loop + NewFamilyInstance** | ~30s | ~3min | ~10min |
| **Loop + Single Transaction** | ~15s | ~1.5min | ~5min |
| **NewFamilyInstances2 + Batch Params** | ~2s | ~10s | ~30s |

#### 10.10.5 Key Implementation Notes

1. **Symbol Activation:** Must activate symbols BEFORE creating `FamilyInstanceCreationData`
2. **Parameter Setting:** Must happen AFTER `NewFamilyInstances2` returns, inside same transaction
3. **Rotation:** Must be applied after creation using `ElementTransformUtils.RotateElement`
4. **Order Preservation:** `NewFamilyInstances2` returns IDs in the same order as input list
5. **Error Handling:** If any creation fails, the entire batch may fail - validate all data first

---

## 11. Success Criteria

### 11.1 Error Handling

- ✅ All critical operations have error handling
- ✅ No unhandled exceptions in production
- ✅ All errors logged with context
- ✅ User-friendly error messages

### 11.2 Data Integrity

- ✅ No partial data corruption
- ✅ All transactions properly managed
- ✅ Rollback works correctly
- ✅ Data consistency maintained

### 11.3 User Experience

- ✅ Clear error messages
- ✅ Progress reporting functional
- ✅ System continues where possible
- ✅ User feedback positive

### 11.4 Performance (NEW)

- ✅ 10x improvement in sleeve placement speed
- ✅ Single transaction per discipline
- ✅ Deferred regeneration enabled
- ✅ BIM 360 sync operations minimized

---

**End of Implementation Plan**

