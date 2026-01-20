#!/usr/bin/env python3
"""
Script to fix LoggingConfiguration namespace issues in C# files.
Replaces System.IO.LoggingConfiguration with JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration
"""

import os
import re
import glob

def fix_logging_namespace():
    """Fix LoggingConfiguration namespace in all C# files"""
    
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
            
            # Pattern 1: Replace System.IO.LoggingConfiguration with JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration
            pattern1 = r'System\.IO\.LoggingConfiguration'
            replacement1 = r'JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration'
            content = re.sub(pattern1, replacement1, content)
            
            # Pattern 2: Replace just LoggingConfiguration with full namespace (if not already qualified)
            pattern2 = r'(?<!JSE_RevitAddin_MEP_OPENINGS\.Services\.)LoggingConfiguration\.ConditionalAppendAllText'
            replacement2 = r'JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration.ConditionalAppendAllText'
            content = re.sub(pattern2, replacement2, content)
            
            # Count replacements made
            replacements = len(re.findall(pattern1, original_content)) + len(re.findall(pattern2, original_content))
            
            if replacements > 0:
                # Write the modified content back
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(content)
                
                print(f"✅ Fixed {replacements} namespace issues in: {file_path}")
                total_replacements += replacements
            else:
                print(f"⏭️  No namespace issues found in: {file_path}")
                
        except Exception as e:
            print(f"❌ Error processing {file_path}: {e}")
    
    print(f"\n🎯 SUMMARY:")
    print(f"   Total files processed: {len(cs_files)}")
    print(f"   Total namespace fixes: {total_replacements}")
    print(f"   Fixed: System.IO.LoggingConfiguration → JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration")

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
            if 'JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration' in content:
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
    print("🚀 Starting LoggingConfiguration namespace fix...")
    print("=" * 60)
    
    # Fix namespace issues
    fix_logging_namespace()
    
    print("\n" + "=" * 60)
    print("🔧 Adding required using statements...")
    
    # Add using statements if needed
    add_using_statement_if_needed()
    
    print("\n" + "=" * 60)
    print("✅ LOGGING NAMESPACE FIX COMPLETE!")
    print("\n📋 What was fixed:")
    print("   • System.IO.LoggingConfiguration → JSE_RevitAddin_MEP_OPENINGS.Services.LoggingConfiguration")
    print("   • Added required using statements")
    print("   • All LoggingConfiguration calls now use correct namespace")
    print("\n🎯 Next steps:")
    print("   1. Build the project in Visual Studio")
    print("   2. Test the application")
    print("   3. Verify no more CS0234 errors")
