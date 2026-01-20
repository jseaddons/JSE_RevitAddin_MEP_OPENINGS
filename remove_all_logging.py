#!/usr/bin/env python3
"""
Script to remove all deployment mode wrappers and DebugLogger calls.
This will clean up the code by removing all logging statements.
"""

import re
import os

def remove_logging_wrappers(file_path):
    """Remove all deployment mode wrappers and DebugLogger calls from a file"""
    with open(file_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    original_lines = lines.copy()
    new_lines = []
    skip_next = False
    i = 0
    
    while i < len(lines):
        line = lines[i]
        
        # Pattern 1: if (!DeploymentConfiguration.DeploymentMode) followed by DebugLogger on next line
        if 'if (!DeploymentConfiguration.DeploymentMode)' in line:
            # Check if next line(s) contain DebugLogger
            j = i + 1
            found_debug_logger = False
            brace_count = 0
            
            # Look ahead to find DebugLogger or opening brace
            while j < len(lines) and j < i + 5:  # Look ahead max 5 lines
                next_line = lines[j]
                if 'DebugLogger.' in next_line:
                    found_debug_logger = True
                    break
                if '{' in next_line:
                    brace_count += next_line.count('{')
                    break
                j += 1
            
            if found_debug_logger:
                # Skip the if statement and the DebugLogger line
                i = j + 1
                continue
            elif brace_count > 0:
                # Skip the if statement and everything until closing brace
                i += 1
                brace_count = 1
                while i < len(lines) and brace_count > 0:
                    if '{' in lines[i]:
                        brace_count += lines[i].count('{')
                    if '}' in lines[i]:
                        brace_count -= lines[i].count('}')
                    i += 1
                continue
        
        # Pattern 2: Standalone DebugLogger calls
        if 'DebugLogger.' in line and not line.strip().startswith('//'):
            i += 1
            continue
        
        # Pattern 3: SafeFileLogger calls (optional - uncomment to remove)
        # if 'SafeFileLogger.' in line and not line.strip().startswith('//'):
        #     i += 1
        #     continue
        
        # Pattern 4: File.AppendAllText calls
        if 'File.AppendAllText' in line and not line.strip().startswith('//'):
            i += 1
            continue
        
        # Pattern 5: Comments about deployment mode
        if '// ✅ DEPLOYMENT MODE:' in line:
            i += 1
            continue
        
        # Keep the line
        new_lines.append(line)
        i += 1
    
    if new_lines != original_lines:
        with open(file_path, 'w', encoding='utf-8') as f:
            f.writelines(new_lines)
        return True, len(original_lines) - len(new_lines)
    return False, 0

def main():
    """Main function"""
    file_path = r'C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Services\FlagManager_Legacy.cs'
    
    if not os.path.exists(file_path):
        print(f"❌ File not found: {file_path}")
        return
    
    print(f"Processing {file_path}...")
    modified, lines_removed = remove_logging_wrappers(file_path)
    
    if modified:
        print(f"✅ Removed {lines_removed} logging lines from {file_path}")
    else:
        print(f"ℹ️ No changes needed in {file_path}")

if __name__ == '__main__':
    main()
