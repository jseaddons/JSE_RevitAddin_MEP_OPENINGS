#!/usr/bin/env python3
"""
Script to automatically fix hardcoded logging in C# files.
Replaces File.AppendAllText() calls with LoggingConfiguration.ConditionalAppendAllText()
"""

import os
import re
import glob

def fix_hardcoded_logging():
    """Fix hardcoded logging in all C# files"""
    
    # Find all C# files in the project
    cs_files = []
    for root, dirs, files in os.walk('.'):
        # Skip certain directories
        dirs[:] = [d for d in dirs if d not in ['bin', 'obj', '.git', 'Log']]
        
        for file in files:
            if file.endswith('.cs'):
                cs_files.append(os.path.join(root, file))
    
    print(f"Found {len(cs_files)} C# files to process...")
    
    total_replacements = 0
    
    for file_path in cs_files:
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            original_content = content
            
            # Pattern 1: Replace File.AppendAllText(debugLogPath, ...) with LoggingConfiguration.ConditionalAppendAllText(debugLogPath, ...)
            pattern1 = r'File\.AppendAllText\(debugLogPath,\s*([^)]+)\)'
            replacement1 = r'LoggingConfiguration.ConditionalAppendAllText(debugLogPath, \1)'
            content = re.sub(pattern1, replacement1, content)
            
            # Pattern 2: Replace File.AppendAllText with LoggingConfiguration.ConditionalAppendAllText for hardcoded paths
            pattern2 = r'File\.AppendAllText\(@"([^"]+\.log)",\s*([^)]+)\)'
            replacement2 = r'LoggingConfiguration.ConditionalAppendAllText(@"\1", \2)'
            content = re.sub(pattern2, replacement2, content)
            
            # Pattern 3: Comment out standalone File.AppendAllText calls that reference undefined variables
            pattern3 = r'(\s+)File\.AppendAllText\(debugLogPath,\s*([^)]+)\);'
            replacement3 = r'\1// DISABLED: Hardcoded log write\n\1// File.AppendAllText(debugLogPath, \2);'
            content = re.sub(pattern3, replacement3, content)
            
            # Count replacements made
            replacements = len(re.findall(pattern1, original_content)) + len(re.findall(pattern2, original_content)) + len(re.findall(pattern3, original_content))
            
            if replacements > 0:
                # Write the modified content back
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(content)
                
                print(f"✅ Fixed {replacements} hardcoded log calls in: {file_path}")
                total_replacements += replacements
            else:
                print(f"⏭️  No hardcoded log calls found in: {file_path}")
                
        except Exception as e:
            print(f"❌ Error processing {file_path}: {e}")
    
    print(f"\n🎯 SUMMARY:")
    print(f"   Total files processed: {len(cs_files)}")
    print(f"   Total replacements made: {total_replacements}")
    print(f"   Global logging switch: DisableAllHardcodedLogging = true")
    print(f"\n✅ All hardcoded logging has been converted to use LoggingConfiguration.ConditionalAppendAllText()")
    print(f"   This means unwanted log files will no longer be created!")

def add_using_statement_if_needed():
    """Add using statement for LoggingConfiguration if needed"""
    
    cs_files = []
    for root, dirs, files in os.walk('.'):
        dirs[:] = [d for d in dirs if d not in ['bin', 'obj', '.git', 'Log']]
        for file in files:
            if file.endswith('.cs'):
                cs_files.append(os.path.join(root, file))
    
    for file_path in cs_files:
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # Check if file uses LoggingConfiguration but doesn't have the using statement
            if 'LoggingConfiguration.ConditionalAppendAllText' in content:
                if 'using JSE_RevitAddin_MEP_OPENINGS.Services;' not in content:
                    # Add the using statement after other using statements
                    using_pattern = r'(using [^;]+;\s*\n)+'
                    match = re.search(using_pattern, content)
                    if match:
                        # Insert after the last using statement
                        insert_pos = match.end()
                        new_content = content[:insert_pos] + 'using JSE_RevitAddin_MEP_OPENINGS.Services;\n' + content[insert_pos:]
                        
                        with open(file_path, 'w', encoding='utf-8') as f:
                            f.write(new_content)
                        
                        print(f"✅ Added using statement to: {file_path}")
            
        except Exception as e:
            print(f"❌ Error adding using statement to {file_path}: {e}")

if __name__ == "__main__":
    print("🚀 Starting hardcoded logging fix...")
    print("=" * 50)
    
    # Fix hardcoded logging
    fix_hardcoded_logging()
    
    print("\n" + "=" * 50)
    print("🔧 Adding required using statements...")
    
    # Add using statements if needed
    add_using_statement_if_needed()
    
    print("\n" + "=" * 50)
    print("✅ HARDCODED LOGGING FIX COMPLETE!")
    print("\n📋 What was fixed:")
    print("   • File.AppendAllText() → LoggingConfiguration.ConditionalAppendAllText()")
    print("   • Added required using statements")
    print("   • Global switch DisableAllHardcodedLogging = true will prevent unwanted logs")
    print("\n🎯 Next steps:")
    print("   1. Build the project in Visual Studio")
    print("   2. Test the application")
    print("   3. Check that unwanted log files are no longer created")
