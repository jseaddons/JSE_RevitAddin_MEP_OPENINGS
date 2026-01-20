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

## 10. Success Criteria

### 10.1 Error Handling

- ✅ All critical operations have error handling
- ✅ No unhandled exceptions in production
- ✅ All errors logged with context
- ✅ User-friendly error messages

### 10.2 Data Integrity

- ✅ No partial data corruption
- ✅ All transactions properly managed
- ✅ Rollback works correctly
- ✅ Data consistency maintained

### 10.3 User Experience

- ✅ Clear error messages
- ✅ Progress reporting functional
- ✅ System continues where possible
- ✅ User feedback positive

---

**End of Implementation Plan**

