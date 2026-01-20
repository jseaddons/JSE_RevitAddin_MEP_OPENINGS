using System;
using System.Collections.Generic;
using System.IO;

class TestConfigParsing
{
    static void Main()
    {
        var content = File.ReadAllText("test_config.txt");
        var lines = content.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
        var currentKey = "";
        var currentValue = "";
        var config = new Dictionary<string, string>();
        
        foreach (var line in lines)
        {
            if (string.IsNullOrWhiteSpace(line))
                continue;
            
            // If line contains "=", it's a new key-value pair
            if (line.Contains("="))
            {
                // Process previous key-value pair if exists
                if (!string.IsNullOrEmpty(currentKey) && !string.IsNullOrEmpty(currentValue))
                {
                    config[currentKey] = currentValue;
                }
                
                // Start new key-value pair
                var parts = line.Split(new char[] { '=' }, 2);
                if (parts.Length == 2)
                {
                    currentKey = parts[0].Trim();
                    currentValue = parts[1].Trim();
                }
            }
            else
            {
                // This is a continuation of the previous value (due to line break)
                if (!string.IsNullOrEmpty(currentValue))
                {
                    currentValue += line.Trim();
                }
            }
        }
        
        // Process the last key-value pair
        if (!string.IsNullOrEmpty(currentKey) && !string.IsNullOrEmpty(currentValue))
        {
            config[currentKey] = currentValue;
        }
        
        // Test the parsing
        if (config.ContainsKey("SelectedReferenceFiles"))
        {
            var files = config["SelectedReferenceFiles"].Split('|');
            Console.WriteLine($"Found {files.Length} reference files:");
            foreach (var file in files)
            {
                Console.WriteLine($"  {file}");
            }
        }
    }
}
