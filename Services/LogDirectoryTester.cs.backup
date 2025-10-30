using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Utility class to test and verify log directory creation
    /// Can be called from a command to verify SafeFileLogger is working
    /// </summary>
    public static class LogDirectoryTester
    {
        /// <summary>
        /// Tests log directory creation and shows a dialog with results
        /// </summary>
        public static void TestLogDirectoryCreation()
        {
            try
            {
                // Force AppData directory creation
                SafeFileLogger.EnsureAppDataLogsDirectory();
                
                // Get current log directory
                string logDir = SafeFileLogger.GetLogDirectory();
                string logDirInfo = SafeFileLogger.GetLogDirectoryInfo();
                
                // Try to write a test file
                string testFileName = "log_directory_test.txt";
                string testMessage = $"Log Directory Test - {DateTime.Now:yyyy-MM-dd HH:mm:ss}\n";
                testMessage += $"Testing SafeFileLogger directory creation and write capability.\n";
                
                SafeFileLogger.SafeAppendText(testFileName, testMessage);
                
                // Check if file was created
                string testFilePath = SafeFileLogger.GetLogFilePath(testFileName);
                bool fileExists = File.Exists(testFilePath);
                string fileContents = fileExists ? File.ReadAllText(testFilePath) : "File not found";
                
                // Show results
                string result = $"═══════════════════════════════════════\n";
                result += $"    LOG DIRECTORY TEST RESULTS\n";
                result += $"═══════════════════════════════════════\n\n";
                result += logDirInfo;
                result += $"\n\n";
                result += $"Test File: {testFileName}\n";
                result += $"Test File Path: {testFilePath}\n";
                result += $"Test File Exists: {fileExists}\n";
                result += $"\n";
                result += $"═══════════════════════════════════════\n";
                result += $"If the AppData Logs directory doesn't exist,\n";
                result += $"SafeFileLogger will use the project Log directory\n";
                result += $"or fall back to Temp/Desktop directories.\n";
                result += $"═══════════════════════════════════════\n";
                
                MessageBox.Show(
                    result,
                    "Log Directory Test",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                
                // Also log to debug output
                System.Diagnostics.Debug.WriteLine(result);
            }
            catch (Exception ex)
            {
                string error = $"Error testing log directory: {ex.Message}\n\n{ex.StackTrace}";
                MessageBox.Show(error, "Log Directory Test Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                System.Diagnostics.Debug.WriteLine(error);
            }
        }
        
        /// <summary>
        /// Gets a simple status message about the log directory
        /// </summary>
        public static string GetStatusMessage()
        {
            try
            {
                string logDir = SafeFileLogger.GetLogDirectory();
                bool exists = Directory.Exists(logDir);
                
                string appDataPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "JSE_MEP_Openings",
                    "Logs"
                );
                bool appDataExists = Directory.Exists(appDataPath);
                
                // Get list of log files in the directory
                string logFiles = "No log files found";
                if (exists)
                {
                    try
                    {
                        var files = Directory.GetFiles(logDir, "*.log", SearchOption.TopDirectoryOnly);
                        if (files.Length > 0)
                        {
                            logFiles = $"{files.Length} log file(s):\n";
                            foreach (var file in files.Take(5)) // Show first 5
                            {
                                logFiles += $"  • {Path.GetFileName(file)}\n";
                            }
                            if (files.Length > 5)
                                logFiles += $"  ... and {files.Length - 5} more";
                        }
                    }
                    catch { }
                }
                
                return $"Current Log Directory: {logDir}\n" +
                       $"Directory Exists: {exists}\n" +
                       $"\n" +
                       $"AppData Logs Path: {appDataPath}\n" +
                       $"AppData Logs Exists: {appDataExists}\n" +
                       $"\n" +
                       $"{logFiles}";
            }
            catch (Exception ex)
            {
                return $"Error getting status: {ex.Message}";
            }
        }
        
        /// <summary>
        /// Opens the log directory in Windows Explorer
        /// </summary>
        public static void OpenLogDirectoryInExplorer()
        {
            try
            {
                string logDir = SafeFileLogger.GetLogDirectory();
                if (Directory.Exists(logDir))
                {
                    System.Diagnostics.Process.Start("explorer.exe", logDir);
                }
                else
                {
                    MessageBox.Show(
                        $"Log directory does not exist: {logDir}\n\n" +
                        $"The directory will be created automatically when you run Refresh.",
                        "Log Directory Not Found",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Error opening log directory: {ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
    }
}
