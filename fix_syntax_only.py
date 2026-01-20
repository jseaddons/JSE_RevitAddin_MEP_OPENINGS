#!/usr/bin/env python3
"""
Script to fix FlagManager_Legacy.cs by only removing duplicate deployment checks.
This is safer than removing all logging - it just fixes the syntax errors.
"""

import re

def fix_file(file_path):
    """Fix duplicate deployment mode checks"""
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    original_content = content
    
    # Fix duplicate if statements on same line
    # Pattern: if (!DeploymentConfiguration.DeploymentMode)     if (!DeploymentConfiguration.DeploymentMode)
    content = re.sub(
        r'if \(!DeploymentConfiguration\.DeploymentMode\)\s+if \(!DeploymentConfiguration\.DeploymentMode\)',
        'if (!DeploymentConfiguration.DeploymentMode)',
        content
    )
    
    # Fix duplicate if statements on consecutive lines
    content = re.sub(
        r'(if \(!DeploymentConfiguration\.DeploymentMode\))\s*\n\s*(if \(!DeploymentConfiguration\.DeploymentMode\))',
        r'\1',
        content
    )
    
    # Fix the specific broken comment pattern from lines 12-14
    content = re.sub(
        r'// Pattern: if \(!DeploymentConfiguration\.DeploymentMode\) \{\s+if \(!DeploymentConfiguration\.DeploymentMode\)\s+DebugLogger\.Info\(\.\.\.\); \}',
        '// Pattern: if (!DeploymentConfiguration.DeploymentMode) { DebugLogger.Info(...); }',
        content
    )
    
    if content != original_content:
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(content)
        return True
    return False

if __name__ == '__main__':
    import sys
    
    # Check if file was provided as argument
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
    else:
        file_path = r'C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Services\FlagManager_Legacy.cs'
    
    print(f"Fixing {file_path}...")
    if fix_file(file_path):
        print(f"✅ Fixed duplicate deployment checks")
    else:
        print(f"ℹ️ No changes needed")
