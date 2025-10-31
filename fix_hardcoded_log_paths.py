#!/usr/bin/env python3
"""
Script to replace hardcoded log file paths with SafeFileLogger.GetLogFilePath()
Pattern: File.AppendAllText(@"C:\...\Log\filename.log", ...)
         OR
         File.AppendAllText(@"C:\...\filename.log", ...)
         
Replaces with:
    string logPath = SafeFileLogger.GetLogFilePath("filename.log");
    File.AppendAllText(logPath, ...);
"""

import os
import re
import glob
import sys

def extract_log_filename(hardcoded_path):
    """Extract the log filename from a hardcoded path"""
    # Extract filename from path like @"C:\...\Log\filename.log" or @"C:\...\filename.log"
    match = re.search(r'[\w\-_]+\.log', hardcoded_path)
    if match:
        return match.group(0)
    return None

def replace_hardcoded_log_paths(file_path):
    """Replace hardcoded log paths with SafeFileLogger.GetLogFilePath()"""
    
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    original_content = content
    changes_count = 0
    
    # Pattern 1: Single line File.AppendAllText(@"C:\...\Log\filename.log", $"message");
    pattern1 = re.compile(
        r'File\.AppendAllText\(@\"[^\"]+\\[^"]+\.log\",\s*\$\"([^\"]+)\"\);',
        re.MULTILINE
    )
    
    def replace_match1(match):
        nonlocal changes_count
        hardcoded_path = match.group(0).split('@"')[1].split('"')[0] if '@"' in match.group(0) else ""
        log_message = match.group(1)
        
        filename = extract_log_filename(hardcoded_path)
        if not filename:
            return match.group(0)  # Can't extract filename, skip
        
        changes_count += 1
        # Generate replacement: declare variable + use it
        var_name = filename.replace('.', '_').replace('-', '_').replace(' ', '_')
        replacement = f'string {var_name}LogPath = SafeFileLogger.GetLogFilePath("{filename}");\n            File.AppendAllText({var_name}LogPath, $"{log_message}");'
        
        return replacement
    
    content = pattern1.sub(replace_match1, content)
    
    # Pattern 2: Multi-line File.AppendAllText(@"C:\...\Log\filename.log", 
    #                                           $"message");
    pattern2 = re.compile(
        r'File\.AppendAllText\(@\"[^\"]+\\[^"]+\.log\",\s*\n\s*\$\"([^\"]+)\"\);',
        re.MULTILINE | re.DOTALL
    )
    
    def replace_match2(match):
        nonlocal changes_count
        hardcoded_path = match.group(0).split('@"')[1].split('"')[0] if '@"' in match.group(0) else ""
        log_message = match.group(1).strip()
        
        filename = extract_log_filename(hardcoded_path)
        if not filename:
            return match.group(0)  # Can't extract filename, skip
        
        changes_count += 1
        var_name = filename.replace('.', '_').replace('-', '_').replace(' ', '_')
        replacement = f'string {var_name}LogPath = SafeFileLogger.GetLogFilePath("{filename}");\n            File.AppendAllText({var_name}LogPath, $"{log_message}");'
        
        return replacement
    
    content = pattern2.sub(replace_match2, content)
    
    # Pattern 3: File.AppendAllText(@"C:\...\Log\filename.log", variable_or_expression);
    pattern3 = re.compile(
        r'File\.AppendAllText\(@\"([^\"]+\\[^"]+\.log)\",\s*([^;]+);',
        re.MULTILINE
    )
    
    def replace_match3(match):
        nonlocal changes_count
        hardcoded_path = match.group(1)
        log_content = match.group(2)
        
        filename = extract_log_filename(hardcoded_path)
        if not filename:
            return match.group(0)  # Can't extract filename, skip
        
        changes_count += 1
        var_name = filename.replace('.', '_').replace('-', '_').replace(' ', '_')
        replacement = f'string {var_name}LogPath = SafeFileLogger.GetLogFilePath("{filename}");\n            File.AppendAllText({var_name}LogPath, {log_content};'
        
        return replacement
    
    content = pattern3.sub(replace_match3, content)
    
    # Pattern 4: File.WriteAllText(@"C:\...\Log\filename.log", ...)
    pattern4 = re.compile(
        r'File\.WriteAllText\(@\"([^\"]+\\[^"]+\.log)\",\s*([^;]+);',
        re.MULTILINE
    )
    
    def replace_match4(match):
        nonlocal changes_count
        hardcoded_path = match.group(1)
        log_content = match.group(2)
        
        filename = extract_log_filename(hardcoded_path)
        if not filename:
            return match.group(0)  # Can't extract filename, skip
        
        changes_count += 1
        var_name = filename.replace('.', '_').replace('-', '_').replace(' ', '_')
        replacement = f'string {var_name}LogPath = SafeFileLogger.GetLogFilePath("{filename}");\n            File.WriteAllText({var_name}LogPath, {log_content};'
        
        return replacement
    
    content = pattern4.sub(replace_match4, content)
    
    # Add using statement if SafeFileLogger is used but not imported
    if 'SafeFileLogger.GetLogFilePath' in content:
        if 'using JSE_RevitAddin_MEP_OPENINGS.Services;' not in content:
            # Find the last using statement
            using_pattern = r'(using [^;]+;\s*\n)'
            matches = list(re.finditer(using_pattern, content))
            if matches:
                last_using = matches[-1]
                insert_pos = last_using.end()
                content = content[:insert_pos] + 'using JSE_RevitAddin_MEP_OPENINGS.Services;\n' + content[insert_pos:]
            else:
                # No using statements, add at the top
                namespace_match = re.search(r'namespace\s+\w+', content)
                if namespace_match:
                    insert_pos = namespace_match.start()
                    content = content[:insert_pos] + 'using JSE_RevitAddin_MEP_OPENINGS.Services;\n' + content[insert_pos:]
    
    if content != original_content:
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(content)
        return changes_count
    
    return 0

def test_single_file(test_file_path=None):
    """Test on a single file first to verify the replacement pattern works"""
    if test_file_path is None:
        # Find an unimportant test file (e.g., in Helpers or Utils)
        test_files = []
        for root, dirs, files in os.walk('.'):
            # Look for files in Helpers or Utils directories (less critical)
            if 'Helper' in root or 'Util' in root:
                for file in files:
                    if file.endswith('.cs') and not file.endswith('.backup.cs'):
                        test_files.append(os.path.join(root, file))
        
        if test_files:
            test_file_path = test_files[0]
        else:
            # Fallback: find any .cs file with hardcoded paths
            for root, dirs, files in os.walk('.'):
                dirs[:] = [d for d in dirs if d not in ['bin', 'obj', '.git', 'Log', 'Scripts', 'Backup']]
                for file in files:
                    if file.endswith('.cs') and not file.endswith('.backup.cs'):
                        file_path = os.path.join(root, file)
                        with open(file_path, 'r', encoding='utf-8') as f:
                            content = f.read()
                        if 'File.AppendAllText(@"C:\\' in content or 'File.WriteAllText(@"C:\\' in content:
                            test_file_path = file_path
                            break
                    if test_file_path:
                        break
                if test_file_path:
                    break
    
    if not test_file_path or not os.path.exists(test_file_path):
        print("⚠️  No test file found with hardcoded paths. Proceeding to process all files...")
        return True  # Continue with full processing
    
    print(f"🧪 TESTING on single file: {test_file_path}\n")
    
    # Backup test file
    backup_path = test_file_path + '.backup'
    if not os.path.exists(backup_path):
        with open(test_file_path, 'r', encoding='utf-8') as f:
            backup_content = f.read()
        with open(backup_path, 'w', encoding='utf-8') as f:
            f.write(backup_content)
        print(f"✅ Backup created: {backup_path}")
    
    # Test replacement
    changes = replace_hardcoded_log_paths(test_file_path)
    
    if changes > 0:
        print(f"\n✅ TEST SUCCESS: Replaced {changes} hardcoded paths in test file")
        print(f"   File: {test_file_path}")
        print(f"\n📋 Please review the changes in the test file.")
        print(f"   If the replacement looks correct, run the script again with --all to process all files.")
        return False  # Don't proceed with all files yet
    else:
        print(f"\n⏭️  No hardcoded paths found in test file.")
        print(f"   Proceeding to process all files...")
        return True  # Continue with full processing

def main():
    # Check if user wants to skip test and process all files
    process_all = '--all' in sys.argv
    
    if not process_all:
        # Test on single file first
        test_file = None
        if len(sys.argv) > 1 and sys.argv[1] != '--all':
            test_file = sys.argv[1]
        
        should_continue = test_single_file(test_file)
        if not should_continue:
            print(f"\n{'='*60}")
            print(f"⚠️  TEST MODE COMPLETE - Run with --all to process all files")
            print(f"{'='*60}")
            return
    
    # Process all files
    print(f"\n{'='*60}")
    print(f"🚀 PROCESSING ALL FILES...")
    print(f"{'='*60}\n")
    
    # Find all .cs files in the project
    cs_files = []
    for root, dirs, files in os.walk('.'):
        # Skip certain directories
        dirs[:] = [d for d in dirs if d not in ['bin', 'obj', '.git', 'Log', 'Scripts', 'Backup']]
        
        for file in files:
            if file.endswith('.cs') and not file.endswith('.backup.cs'):
                cs_files.append(os.path.join(root, file))
    
    print(f"Found {len(cs_files)} C# files to process...\n")
    
    total_changes = 0
    files_changed = []
    
    for file_path in sorted(cs_files):
        try:
            changes = replace_hardcoded_log_paths(file_path)
            if changes > 0:
                total_changes += changes
                files_changed.append((file_path, changes))
                print(f"✅ {file_path}: {changes} replacements")
        except Exception as e:
            print(f"❌ Error processing {file_path}: {e}")
    
    print(f"\n{'='*60}")
    print(f"SUMMARY:")
    print(f"   Total files processed: {len(cs_files)}")
    print(f"   Total replacements made: {total_changes}")
    print(f"   Files modified: {len(files_changed)}")
    if files_changed:
        print(f"\nModified files:")
        for file_path, count in files_changed[:10]:  # Show first 10
            print(f"   - {file_path} ({count} changes)")
        if len(files_changed) > 10:
            print(f"   ... and {len(files_changed) - 10} more files")
    print(f"{'='*60}")
    print(f"\n✅ Hardcoded log paths replaced with SafeFileLogger.GetLogFilePath()")

if __name__ == "__main__":
    main()

