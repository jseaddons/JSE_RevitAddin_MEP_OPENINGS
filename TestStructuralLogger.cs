using System;
using System.IO;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS
{
    /// <summary>
    /// Test class to verify StructuralElementLogger is working
    /// </summary>
    public class TestStructuralLogger
    {
        public static void Main(string[] args)
        {
            Console.WriteLine("Testing StructuralElementLogger...");
            
            try
            {
                // Initialize the logger
                StructuralElementLogger.InitializeLogger();
                Console.WriteLine($"Logger initialized. Status: {(StructuralElementLogger.IsLoggerInitialized() ? "SUCCESS" : "FAILED")}");
                Console.WriteLine($"Log file path: {StructuralElementLogger.GetLogFilePath()}");
                
                // Test logging some entries
                StructuralElementLogger.LogStructuralElement("TEST_ELEMENT", new Autodesk.Revit.DB.ElementId(12345L), "TEST_DETECTED", "This is a test entry");
                StructuralElementLogger.LogIntersection("TestDuct", new Autodesk.Revit.DB.ElementId(111L), "TestWall", new Autodesk.Revit.DB.ElementId(222L), "Test intersection point");
                StructuralElementLogger.LogSleevePlace("TestWall", new Autodesk.Revit.DB.ElementId(222L), new Autodesk.Revit.DB.ElementId(333L), "Test placement point", "100x200");
                StructuralElementLogger.LogSummary("TestCommand", 10, 5, 3, 2);
                
                Console.WriteLine("Test logging completed.");
                
                // Check if file exists and show content
                string logPath = StructuralElementLogger.GetLogFilePath();
                if (File.Exists(logPath))
                {
                    Console.WriteLine("\n=== LOG FILE CONTENT ===");
                    string content = File.ReadAllText(logPath);
                    Console.WriteLine(content);
                    Console.WriteLine("=== END LOG FILE ===\n");
                }
                else
                {
                    Console.WriteLine("ERROR: Log file does not exist!");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ERROR: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
            }
            
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
    }
}
