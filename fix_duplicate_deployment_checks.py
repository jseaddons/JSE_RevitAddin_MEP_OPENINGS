#!/usr/bin/env python3
"""
Script to remove duplicate if (!DeploymentConfiguration.DeploymentMode) checks
that were added by the wrap_logging_for_deployment.py script.
"""

import re

def fix_duplicate_deployment_checks(file_path):
    """Remove duplicate deployment mode checks from a file"""
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    original_content = content
    
    # Pattern 1: Duplicate on same line (nested)
    # if (!DeploymentConfiguration.DeploymentMode)                if (!DeploymentConfiguration.DeploymentMode)
    pattern1 = r'if \(!DeploymentConfiguration\.DeploymentMode\)\s+if \(!DeploymentConfiguration\.DeploymentMode\)'
    content = re.sub(pattern1, 'if (!DeploymentConfiguration.DeploymentMode)', content)
    
    # Pattern 2: Duplicate on consecutive lines with varying indentation
    # This matches cases where the second if statement is on the next line
    pattern2 = r'(if \(!DeploymentConfiguration\.DeploymentMode\))\s*\n\s*(if \(!DeploymentConfiguration\.DeploymentMode\))'
    content = re.sub(pattern2, r'\1', content)
    
    if content != original_content:
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(content)
        return True
    return False

if __name__ == '__main__':
    file_path = r'C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Services\FlagManager_Legacy.cs'
    if fix_duplicate_deployment_checks(file_path):
        print(f"✅ Fixed duplicate deployment checks in {file_path}")
    else:
        print(f"ℹ️ No changes needed in {file_path}")
