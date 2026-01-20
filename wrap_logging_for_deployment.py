#!/usr/bin/env python3
"""
Script to wrap all logging calls with DeploymentConfiguration.DeploymentMode checks.
Tests pattern on FlagManager first, then applies to all service classes.
"""

import re
import os
from pathlib import Path

# Pattern to find DebugLogger calls
DEBUG_LOGGER_PATTERN = r'DebugLogger\.(Info|Error|Warning|Debug|Critical|Log)\s*\(([^;]+)\);'

# Pattern to find File.AppendAllText calls
FILE_APPEND_PATTERN = r'File\.AppendAllText\s*\(([^)]+)\);'

# Pattern to find SafeFileLogger calls (these are already safe, but check)
SAFE_FILE_LOGGER_PATTERN = r'SafeFileLogger\.\w+\(([^)]+)\);'

def wrap_debug_logger_call(match):
    """Wrap a DebugLogger call with deployment mode check"""
    method = match.group(1)
    args = match.group(2)
    # Indent the original call
    wrapped = f'''if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.{method}({args});'''
    return wrapped

def wrap_file_append_call(match):
    """Wrap a File.AppendAllText call with deployment mode check"""
    args = match.group(1)
    wrapped = f'''// ✅ DEPLOYMENT MODE: Skip file writes
                            if (!DeploymentConfiguration.DeploymentMode)
                            {{
                                File.AppendAllText({args});
                            }}'''
    return wrapped

def process_file(file_path):
    """Process a single C# file to wrap logging calls"""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        original_content = content
        
        # Check if DeploymentConfiguration is already imported
        if 'DeploymentConfiguration' not in content:
            # Find the last using statement
            using_pattern = r'(using\s+[^;]+;)'
            using_statements = re.findall(using_pattern, content)
            if using_statements:
                last_using = using_statements[-1]
                content = content.replace(last_using, 
                    f'{last_using}\nusing JSE_RevitAddin_MEP_OPENINGS.Services; // For DeploymentConfiguration')
        
        # Wrap DebugLogger calls that aren't already wrapped
        def replace_debug_logger(m):
            full_match = m.group(0)
            # Check if already wrapped
            if 'DeploymentConfiguration.DeploymentMode' in full_match:
                return full_match  # Already wrapped, skip
            
            method = m.group(1)
            args = m.group(2)
            # Get indentation from the line
            lines_before = content[:m.start()].split('\n')
            indent = len(lines_before[-1]) - len(lines_before[-1].lstrip())
            indent_str = ' ' * indent
            
            wrapped = f'''{indent_str}if (!DeploymentConfiguration.DeploymentMode)
{indent_str}    DebugLogger.{method}({args});'''
            return wrapped
        
        content = re.sub(DEBUG_LOGGER_PATTERN, replace_debug_logger, content)
        
        # Wrap File.AppendAllText calls that aren't already wrapped
        def replace_file_append(m):
            full_match = m.group(0)
            # Check if already wrapped
            if 'DeploymentConfiguration.DeploymentMode' in full_match:
                return full_match  # Already wrapped, skip
            
            args = m.group(1)
            # Get indentation
            lines_before = content[:m.start()].split('\n')
            indent = len(lines_before[-1]) - len(lines_before[-1].lstrip())
            indent_str = ' ' * indent
            
            wrapped = f'''{indent_str}// ✅ DEPLOYMENT MODE: Skip file writes
{indent_str}if (!DeploymentConfiguration.DeploymentMode)
{indent_str}{{
{indent_str}    File.AppendAllText({args});
{indent_str}}}'''
            return wrapped
        
        content = re.sub(FILE_APPEND_PATTERN, replace_file_append, content)
        
        if content != original_content:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(content)
            return True
        return False
    except Exception as e:
        print(f"Error processing {file_path}: {e}")
        return False

def main():
    """Main function"""
    services_dir = Path('Services')
    
    # First test on FlagManager (already done manually)
    print("[OK] FlagManager.cs already wrapped as test case")
    
    # Find all C# service files
    service_files = list(services_dir.glob('*.cs'))
    
    # Exclude FlagManager (already done) and DeploymentConfiguration itself
    files_to_process = [f for f in service_files 
                       if f.name != 'FlagManager.cs' 
                       and f.name != 'DeploymentConfiguration.cs'
                       and f.name != 'DebugLogger.cs']  # DebugLogger handles its own deployment mode
    
    print(f"\nFound {len(files_to_process)} service files to process")
    
    modified_count = 0
    for file_path in files_to_process:
        print(f"Processing {file_path.name}...", end=' ')
        if process_file(file_path):
            print("[MODIFIED]")
            modified_count += 1
        else:
            print("[NO CHANGES]")
    
    print(f"\n[COMPLETE] {modified_count} files modified")

if __name__ == '__main__':
    main()

