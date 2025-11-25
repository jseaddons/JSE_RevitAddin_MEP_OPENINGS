#!/usr/bin/env python3
"""
Script to replace System.IO.File.AppendAllText with SafeFileLogger.SafeAppendText
in UniversalClusterService.cs

This script:
1. Finds all File.AppendAllText calls
2. Extracts the log file path variable
3. Maps common log file paths to their filenames
4. Replaces with SafeFileLogger.SafeAppendText(filename, message)
"""

import re
import sys
from pathlib import Path

# Map of common log path variable names to their log file names
LOG_FILE_MAPPING = {
    'clusterDebugLogPath': 'cluster_debug.log',
    'clusterLogPath': 'cluster_debug.log',
    'placementDebugPath': 'placement_debug.log',
    'flagStateDebugLogPath': 'flag_state_debug.log',
    'orchestratorDebugLogPath': 'orchestrator_debug.log',
    'cluster_bbox_detailed.log': 'cluster_bbox_detailed.log',
}

def extract_log_filename(path_var, line):
    """Extract log filename from path variable or string literal."""
    # Check if it's a direct string literal
    if path_var.startswith('"') or path_var.startswith("'"):
        # Extract filename from string literal
        match = re.search(r'["\']([^"\']+\.log)["\']', path_var)
        if match:
            return match.group(1)
    
    # Check if it's a variable that we know the mapping for
    var_name = path_var.strip()
    if var_name in LOG_FILE_MAPPING:
        return LOG_FILE_MAPPING[var_name]
    
    # Try to find SafeFileLogger.GetLogFilePath call on previous lines
    # This is a fallback - we'll use the variable name as a hint
    if 'cluster' in var_name.lower():
        return 'cluster_debug.log'
    elif 'placement' in var_name.lower():
        return 'placement_debug.log'
    elif 'flag' in var_name.lower():
        return 'flag_state_debug.log'
    elif 'orchestrator' in var_name.lower():
        return 'orchestrator_debug.log'
    
    # Default fallback
    return 'cluster_debug.log'

def replace_file_appendalltext(content):
    """Replace all File.AppendAllText calls with SafeFileLogger.SafeAppendText."""
    
    # Pattern 1: System.IO.File.AppendAllText(pathVar, message)
    pattern1 = re.compile(
        r'System\.IO\.File\.AppendAllText\s*\(\s*([^,]+)\s*,\s*(.+?)\s*\)',
        re.DOTALL
    )
    
    # Pattern 2: File.AppendAllText(pathVar, message) - without System.IO prefix
    pattern2 = re.compile(
        r'File\.AppendAllText\s*\(\s*([^,]+)\s*,\s*(.+?)\s*\)',
        re.DOTALL
    )
    
    def replace_match(match):
        path_var = match.group(1).strip()
        message = match.group(2).strip()
        
        # Extract log filename
        log_filename = extract_log_filename(path_var, match.group(0))
        
        # Handle multi-line messages (remove trailing newlines/whitespace from message)
        message = message.rstrip()
        
        # Return replacement
        return f'SafeFileLogger.SafeAppendText("{log_filename}", {message})'
    
    # Apply both patterns
    content = pattern1.sub(replace_match, content)
    content = pattern2.sub(replace_match, content)
    
    # Also handle cases where the path variable is defined on a previous line
    # and we need to remove the variable declaration if it's only used for logging
    # This is more complex, so we'll do a simpler pass first
    
    return content

def remove_unused_log_path_variables(content):
    """Remove log path variable declarations that are no longer needed."""
    lines = content.split('\n')
    new_lines = []
    i = 0
    
    while i < len(lines):
        line = lines[i]
        
        # Check if this line declares a log path variable using SafeFileLogger.GetLogFilePath
        # and the next line(s) don't use it (because we replaced File.AppendAllText)
        match = re.match(r'(\s*)string\s+(\w+LogPath)\s*=\s*SafeFileLogger\.GetLogFilePath\(["\']([^"\']+)["\']\)\s*;', line)
        
        if match:
            indent = match.group(1)
            var_name = match.group(2)
            log_file = match.group(3)
            
            # Check if this variable is used in the next few lines
            # (we already replaced File.AppendAllText, so it shouldn't be used)
            used = False
            for j in range(i + 1, min(i + 10, len(lines))):
                if var_name in lines[j] and 'SafeFileLogger.SafeAppendText' not in lines[j]:
                    # Variable is still used somewhere
                    used = True
                    break
            
            if not used:
                # Variable is no longer needed, skip this line
                i += 1
                continue
        
        new_lines.append(line)
        i += 1
    
    return '\n'.join(new_lines)

def main():
    if len(sys.argv) > 1:
        file_path = Path(sys.argv[1])
    else:
        file_path = Path('Services/UniversalClusterService.cs')
    
    if not file_path.exists():
        print(f"Error: {file_path} not found!")
        sys.exit(1)
    
    print(f"Reading {file_path}...")
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    original_content = content
    
    # Count occurrences before replacement
    count_before = len(re.findall(r'System\.IO\.File\.AppendAllText|File\.AppendAllText', content))
    print(f"Found {count_before} File.AppendAllText calls")
    
    # Replace File.AppendAllText calls
    print("Replacing File.AppendAllText calls...")
    content = replace_file_appendalltext(content)
    
    # Count occurrences after replacement
    count_after = len(re.findall(r'System\.IO\.File\.AppendAllText|File\.AppendAllText', content))
    print(f"Remaining File.AppendAllText calls: {count_after}")
    
    # Remove unused log path variable declarations
    print("Removing unused log path variable declarations...")
    content = remove_unused_log_path_variables(content)
    
    # Write back
    if content != original_content:
        # Create backup
        backup_path = file_path.with_suffix('.cs.backup_before_logging_replace')
        print(f"Creating backup: {backup_path}")
        with open(backup_path, 'w', encoding='utf-8') as f:
            f.write(original_content)
        
        print(f"Writing updated content to {file_path}...")
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(content)
        
        print(f"✅ Successfully replaced {count_before - count_after} File.AppendAllText calls!")
    else:
        print("No changes made.")

if __name__ == '__main__':
    main()

