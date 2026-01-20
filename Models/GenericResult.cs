using System;

namespace JSE_RevitAddin_MEP_OPENINGS.Models
{
    public class GenericResult<T>
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public T Data { get; set; }

        public static GenericResult<T> Ok(T data, string message = null)
        {
            return new GenericResult<T> { Success = true, Data = data, Message = message };
        }

        public static GenericResult<T> Fail(string message)
        {
            return new GenericResult<T> { Success = false, Message = message };
        }
    }
}
