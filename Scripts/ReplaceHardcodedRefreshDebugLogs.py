"""
Script to replace hardcoded File.AppendAllText calls to refresh_debug.log 
with DebugLogger.Info calls so they respect DeploymentMode flag.
"""

import re
import os

def replace_hardcoded_logs(file_path):
    """Replace File.AppendAllText calls with DebugLogger.Info calls"""
    
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    original_content = content
    changes_count = 0
    
    # Pattern 1: File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\refresh_debug.log", 
    #                                $"...");
    pattern1 = re.compile(
        r'File\.AppendAllText\(@"C:\\JSE_CSharp_Projects\\JSE_MEPOPENING_23\\Log\\refresh_debug\.log",\s*\n\s*\$"([^"]+)"\);',
        re.MULTILINE
    )
    
    def replace_match1(match):
        nonlocal changes_count
        changes_count += 1
        log_message = match.group(1)
        # Escape any newlines and quotes in the message
        log_message = log_message.replace('\n', '\\n').replace('"', '\\"')
        return f'DebugLogger.Info($"[CLASH_DEBUG] {log_message}");'
    
    content = pattern1.sub(replace_match1, content)
    
    # Pattern 2: Single line File.AppendAllText(@"C:\...\refresh_debug.log", $"...");
    pattern2 = re.compile(
        r'File\.AppendAllText\(@"C:\\JSE_CSharp_Projects\\JSE_MEPOPENING_23\\Log\\refresh_debug\.log",\s*\$"([^"]+)"\);',
        re.MULTILINE
    )
    
    def replace_match2(match):
        nonlocal changes_count
        changes_count += 1
        log_message = match.group(1)
        # Remove the [CLASH_DEBUG] prefix if already present to avoid duplication
        if '[CLASH_DEBUG]' in log_message:
            clean_message = log_message.replace('[CLASH_DEBUG]', '').strip()
        else:
            clean_message = log_message
        # Escape any quotes
        clean_message = clean_message.replace('"', '\\"')
        return f'DebugLogger.Info($"[CLASH_DEBUG] {clean_message}");'
    
    content = pattern2.sub(replace_match2, content)
    
    # Pattern 3: Multi-line with string concatenation
    pattern3 = re.compile(
        r'File\.AppendAllText\(@"C:\\JSE_CSharp_Projects\\JSE_MEPOPENING_23\\Log\\refresh_debug\.log",\s*\n\s*\$"([^"]+)"\s*\+\s*\$"([^"]+)"\);',
        re.MULTILINE | re.DOTALL
    )
    
    def replace_match3(match):
        nonlocal changes_count
        changes_count += 1
        msg1 = match.group(1).replace('\\n', '\n')
        msg2 = match.group(2).replace('\\n', '\n')
        log_message = msg1 + msg2
        # Remove [CLASH_DEBUG] if present
        if '[CLASH_DEBUG]' in log_message:
            log_message = log_message.replace('[CLASH_DEBUG]', '').strip()
        log_message = log_message.replace('"', '\\"')
        return f'DebugLogger.Info($"[CLASH_DEBUG] {log_message}");'
    
    content = pattern3.sub(replace_match3, content)
    
    if content != original_content:
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(content)
        return changes_count
    
    return 0

def main():
    # Find RefreshService.cs
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.join(script_dir, '..')
    refresh_service_path = os.path.join(project_root, 'Services', 'RefreshService.cs')
    
    if not os.path.exists(refresh_service_path):
        print(f"Error: Could not find {refresh_service_path}")
        return
    
    print(f"Processing: {refresh_service_path}")
    
    # Backup original file
    backup_path = refresh_service_path + '.backup'
    with open(refresh_service_path, 'r', encoding='utf-8') as f:
        backup_content = f.read()
    
    with open(backup_path, 'w', encoding='utf-8') as f:
        f.write(backup_content)
    print(f"Backup created: {backup_path}")
    
    # Replace hardcoded logs
    changes = replace_hardcoded_logs(refresh_service_path)
    
    if changes > 0:
        print(f"✅ Successfully replaced {changes} hardcoded File.AppendAllText calls with DebugLogger.Info")
        print(f"✅ All logging now respects DeploymentMode flag")
    else:
        print("No changes needed - no hardcoded refresh_debug.log calls found")

if __name__ == '__main__':
    main()

