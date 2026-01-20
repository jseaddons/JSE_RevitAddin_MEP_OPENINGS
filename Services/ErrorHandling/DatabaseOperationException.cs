using System;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.ErrorHandling
{
    /// <summary>
    /// ✅ OOP ERROR HANDLING: Custom exception for database operations
    /// </summary>
    public class DatabaseOperationException : Exception
    {
        public string TableName { get; }
        public string Operation { get; }
        public Dictionary<string, object> Parameters { get; }

        public DatabaseOperationException(
            string tableName,
            string operation,
            string message,
            Exception innerException = null,
            Dictionary<string, object> parameters = null)
            : base(message, innerException)
        {
            TableName = tableName;
            Operation = operation;
            Parameters = parameters ?? new Dictionary<string, object>();
        }
    }

    /// <summary>
    /// Exception for filter operations
    /// </summary>
    public class FilterOperationException : DatabaseOperationException
    {
        public string FilterName { get; }
        public string Category { get; }

        public FilterOperationException(
            string filterName,
            string category,
            string operation,
            string message,
            Exception innerException = null)
            : base("Filters", operation, message, innerException)
        {
            FilterName = filterName;
            Category = category;
        }
    }

    /// <summary>
    /// Exception for file combo operations
    /// </summary>
    public class FileComboOperationException : DatabaseOperationException
    {
        public int ComboId { get; }
        public int FilterId { get; }

        public FileComboOperationException(
            int comboId,
            int filterId,
            string operation,
            string message,
            Exception innerException = null)
            : base("FileCombos", operation, message, innerException)
        {
            ComboId = comboId;
            FilterId = filterId;
        }
    }

    /// <summary>
    /// Exception for clash zone operations
    /// </summary>
    public class ClashZoneOperationException : DatabaseOperationException
    {
        public Guid ClashZoneGuid { get; }

        public ClashZoneOperationException(
            Guid clashZoneGuid,
            string operation,
            string message,
            Exception innerException = null)
            : base("ClashZones", operation, message, innerException)
        {
            ClashZoneGuid = clashZoneGuid;
        }
    }
}


