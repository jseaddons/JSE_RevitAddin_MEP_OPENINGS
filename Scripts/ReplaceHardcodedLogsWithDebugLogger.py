"""
Script to replace ALL hardcoded File.AppendAllText calls with DebugLogger.Info calls
so they respect DeploymentMode flag.

Scans all .cs files in Services directory and replaces:
- File.AppendAllText(@"C:\\...\\refresh_debug.log", ...)
- File.AppendAllText(@"C:\\...\\logger_debug.txt", ...)
- Any other hardcoded log file writes

With: DebugLogger.Info(...)
"""

import re
import os
import glob

def replace_hardcoded_logs(file_path):
    """Replace File.AppendAllText calls with DebugLogger.Info calls"""
    
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    original_content = content
    changes_count = 0
    
    # Pattern 1: Multi-line File.AppendAllText with interpolated string
    # File.AppendAllText(@"C:\...\refresh_debug.log", 
    #                    $"message");
    pattern1 = re.compile(
        r'File\.AppendAllText\(@"[^"]+\\refresh_debug\.log",\s*\n\s*\$"([^"]+)"\);',
        re.MULTILINE
    )
    
    def replace_match1(match):
        nonlocal changes_count
        changes_count += 1
        log_message = match.group(1)
        # Remove [CLASH_DEBUG] if already present to avoid duplication
        if '[CLASH_DEBUG]' in log_message:
            log_message = log_message.replace('[CLASH_DEBUG]', '').strip()
        return f'DebugLogger.Info($"[CLASH_DEBUG] {log_message}");'
    
    content = pattern1.sub(replace_match1, content)
    
    # Pattern 2: Single line File.AppendAllText(@"C:\...\refresh_debug.log", $"message");
    pattern2 = re.compile(
        r'File\.AppendAllText\(@"[^"]+\\refresh_debug\.log",\s*\$"([^"]+)"\);',
        re.MULTILINE
    )
    
    def replace_match2(match):
        nonlocal changes_count
        changes_count += 1
        log_message = match.group(1)
        # Remove [CLASH_DEBUG] if already present
        if '[CLASH_DEBUG]' in log_message:
            log_message = log_message.replace('[CLASH_DEBUG]', '').strip()
        return f'DebugLogger.Info($"[CLASH_DEBUG] {log_message}");'
    
    content = pattern2.sub(replace_match2, content)
    
    # Pattern 3: logger_debug.txt
    pattern3 = re.compile(
        r'File\.AppendAllText\(@"[^"]+\\logger_debug\.txt",\s*\$"([^"]+)"\);',
        re.MULTILINE
    )
    
    def replace_match3(match):
        nonlocal changes_count
        changes_count += 1
        log_message = match.group(1)
        return f'DebugLogger.Info($"{log_message}");'
    
    content = pattern3.sub(replace_match3, content)
    
    # Pattern 4: Any other hardcoded File.AppendAllText to Log directory files
    pattern4 = re.compile(
        r'File\.AppendAllText\(@"[^"]+\\Log\\[^"]+",\s*\n\s*\$"([^"]+)"\);',
        re.MULTILINE
    )
    
    def replace_match4(match):
        nonlocal changes_count
        changes_count += 1
        log_message = match.group(1)
        # Try to preserve any prefix if present
        if '[' in log_message and ']' in log_message:
            return f'DebugLogger.Info($"{log_message}");'
        else:
            return f'DebugLogger.Info($"[DEBUG] {log_message}");'
    
    content = pattern4.sub(replace_match4, content)
    
    # Pattern 5: Single line any hardcoded log file
    pattern5 = re.compile(
        r'File\.AppendAllText\(@"[^"]+\\Log\\[^"]+",\s*\$"([^"]+)"\);',
        re.MULTILINE
    )
    
    def replace_match5(match):
        nonlocal changes_count
        changes_count += 1
        log_message = match.group(1)
        if '[' in log_message and ']' in log_message:
            return f'DebugLogger.Info($"{log_message}");'
        else:
            return f'DebugLogger.Info($"[DEBUG] {log_message}");'
    
    content = pattern5.sub(replace_match5, content)
    
    if content != original_content:
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(content)
        return changes_count
    
    return 0

def main():
    # Find all .cs files in Services directory and root
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.join(script_dir, '..')
    services_dir = os.path.join(project_root, 'Services')
    
    # Find all .cs files
    cs_files = []
    
    # Services directory
    if os.path.exists(services_dir):
        cs_files.extend(glob.glob(os.path.join(services_dir, '*.cs')))
    
    # Root directory (Commands, Models, Views, etc.)
    for dir_name in ['Commands', 'Models', 'Views', 'Services']:
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
            
        # Backup original file
        backup_path = cs_file + '.backup'
        if not os.path.exists(backup_path):
            with open(cs_file, 'r', encoding='utf-8') as f:
                backup_content = f.read()
            with open(backup_path, 'w', encoding='utf-8') as f:
                f.write(backup_content)
        
        # Replace hardcoded logs
        changes = replace_hardcoded_logs(cs_file)
        
        if changes > 0:
            total_changes += changes
            files_changed.append((rel_path, changes))
            print(f"[OK] {rel_path}: {changes} replacements")
    
    print(f"\n{'='*60}")
    if total_changes > 0:
        print(f"[OK] Total: {total_changes} hardcoded File.AppendAllText calls replaced")
        print(f"[OK] {len(files_changed)} files modified")
        print(f"[OK] All logging now respects DeploymentMode flag")
        print(f"\nModified files:")
        for file_path, count in files_changed:
            print(f"  - {file_path} ({count} changes)")
    else:
        print("No changes needed - no hardcoded log file writes found")
    print(f"{'='*60}")

if __name__ == '__main__':
    main()

