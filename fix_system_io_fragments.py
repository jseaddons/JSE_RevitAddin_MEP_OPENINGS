#!/usr/bin/env python3
"""
Fix script-generated System.IO. fragments
"""
import re
from pathlib import Path

def fix_file(file_path):
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    original = content
    
    # Remove standalone "System.IO.                    //" lines
    content = re.sub(r'^\s*System\.IO\.\s+//.*$\n', '', content, flags=re.MULTILINE)
    
    # Fix nested deployment mode checks
    content = re.sub(
        r'if \(!DeploymentConfiguration\.DeploymentMode\)\s*\{\s*// ✅ DEPLOYMENT MODE.*?\n\s*if \(!DeploymentConfiguration\.DeploymentMode\)\s*\{',
        'if (!DeploymentConfiguration.DeploymentMode) {',
        content,
        flags=re.DOTALL
    )
    
    if content != original:
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(content)
        return True
    return False

# Fix all files
services_dir = Path('Services')
modified = 0
for file in services_dir.glob('*.cs'):
    if fix_file(file):
        modified += 1
        print(f"Fixed: {file.name}")

print(f"\nFixed {modified} files")

