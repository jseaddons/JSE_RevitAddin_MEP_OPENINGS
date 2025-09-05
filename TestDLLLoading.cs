using System;
using System.IO;

namespace JSE_RevitAddin_MEP_OPENINGS
{
    /// <summary>
    /// Static test class to verify DLL loading without any Revit dependencies
    /// </summary>
    public static class TestDLLLoading
    {
        /// <summary>
        /// Basic test to see if DLL can be loaded and executed
        /// </summary>
        public static void BasicTest()
        {
            try
            {
                string testLogPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\dll_load_test.log";
                File.AppendAllText(testLogPath, $"[{DateTime.Now}] DLL LOADED SUCCESSFULLY - BasicTest() called\n");
                File.AppendAllText(testLogPath, $"[{DateTime.Now}] Assembly: {System.Reflection.Assembly.GetExecutingAssembly().Location}\n");
                File.AppendAllText(testLogPath, $"[{DateTime.Now}] DateTime: {DateTime.Now}\n");
                File.AppendAllText(testLogPath, $"[{DateTime.Now}] DLL is functional and can write files\n");
            }
            catch (Exception ex)
            {
                try
                {
                    File.AppendAllText(@"C:\temp\dll_load_error.log", $"[{DateTime.Now}] BasicTest failed: {ex.Message}\n{ex.StackTrace}\n");
                }
                catch
                {
                    // If we can't even write to temp, there's a serious problem
                }
            }
        }

        /// <summary>
        /// Test with forced exception handling to verify error logging
        /// </summary>
        public static void TestWithError()
        {
            try
            {
                string testLogPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\dll_error_test.log";
                File.AppendAllText(testLogPath, $"[{DateTime.Now}] TestWithError() called\n");
                
                // Deliberately cause an error to test error handling
                object nullObject = null;
                var test = nullObject.ToString(); // This will throw NullReferenceException
            }
            catch (Exception ex)
            {
                try
                {
                    string errorLogPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\dll_error_caught.log";
                    File.AppendAllText(errorLogPath, $"[{DateTime.Now}] Expected error caught: {ex.Message}\n");
                    File.AppendAllText(errorLogPath, $"[{DateTime.Now}] Error handling works correctly\n");
                }
                catch (Exception ex2)
                {
                    File.AppendAllText(@"C:\temp\dll_error_fallback.log", $"[{DateTime.Now}] Error handling failed: {ex2.Message}\n");
                }
            }
        }

        /// <summary>
        /// Test static constructor to see if it runs when DLL is loaded
        /// </summary>
        static TestDLLLoading()
        {
            try
            {
                string ctorLogPath = @"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\dll_static_constructor.log";
                File.AppendAllText(ctorLogPath, $"[{DateTime.Now}] TestDLLLoading STATIC CONSTRUCTOR called\n");
                File.AppendAllText(ctorLogPath, $"[{DateTime.Now}] This runs when the class is first accessed\n");
            }
            catch (Exception ex)
            {
                try
                {
                    File.AppendAllText(@"C:\temp\dll_static_ctor_error.log", $"[{DateTime.Now}] Static constructor failed: {ex.Message}\n");
                }
                catch { }
            }
        }
    }
}
