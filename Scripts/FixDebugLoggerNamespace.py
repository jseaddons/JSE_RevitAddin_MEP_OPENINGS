"""
Script to fix incorrect DebugLogger namespace after ReplaceHardcodedLogsWithDebugLogger.py
Replaces: System.IO.DebugLogger.Info
With: DebugLogger.Info
"""

import re
import os
import glob

def fix_debuglogger_namespace(file_path):
    """Replace System.IO.DebugLogger.Info with DebugLogger.Info"""
    
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    original_content = content
    changes_count = 0
    
    # Replace System.IO.DebugLogger.Info with DebugLogger.Info
    pattern = re.compile(r'System\.IO\.DebugLogger\.Info')
    
    def replace_match(match):
        nonlocal changes_count
        changes_count += 1
        return 'DebugLogger.Info'
    
    content = pattern.sub(replace_match, content)
    
    if content != original_content:
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(content)
        return changes_count
    
    return 0

def main():
    # Find all .cs files in Services, Commands, Models, Views directories
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.join(script_dir, '..')
    
    cs_files = []
    
    # Find all .cs files
    for dir_name in ['Services', 'Commands', 'Models', 'Views']:
        dir_path = os.path.join(project_root, dir_name)
        if os.path.exists(dir_path):
            cs_files.extend(glob.glob(os.path.join(dir_path, '*.cs')))
    
    # Remove duplicates
    cs_files = list(set(cs_files))
    
    if not cs_files:
        print(f"Error: Could not find any .cs files in {project_root}")
        return
    
    print(f"Found {len(cs_files)} .cs files to process\n")
    
    total_changes = 0
    files_changed = []
    
    for cs_file in sorted(cs_files):
        rel_path = os.path.relpath(cs_file, project_root)
        
        # Skip backup files
        if '.backup' in cs_file:
            continue
        
        # Fix DebugLogger namespace
        changes = fix_debuglogger_namespace(cs_file)
        
        if changes > 0:
            total_changes += changes
            files_changed.append((rel_path, changes))
            print(f"[OK] {rel_path}: {changes} replacements")
    
    print(f"\n{'='*60}")
    if total_changes > 0:
        print(f"[OK] Total: {total_changes} System.IO.DebugLogger.Info calls fixed")
        print(f"[OK] {len(files_changed)} files modified")
        print(f"\nModified files:")
        for file_path, count in files_changed:
            print(f"  - {file_path} ({count} changes)")
    else:
        print("No changes needed - no System.IO.DebugLogger.Info found")
    print(f"{'='*60}")

if __name__ == '__main__':
    main()

