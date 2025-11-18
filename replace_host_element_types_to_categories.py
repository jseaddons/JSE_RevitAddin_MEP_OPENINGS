#!/usr/bin/env python3
"""
Script to replace SelectedHostElementTypes with SelectedHostCategories throughout the codebase.
Step 1: Delete redundant SelectedHostElementTypes code where SelectedHostCategories already exists
Step 2: Replace remaining SelectedHostElementTypes with SelectedHostCategories
"""

import os
import re
from pathlib import Path
from typing import List, Tuple, Dict

# Define the workspace root
WORKSPACE_ROOT = Path(r"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23")

# Files to process (C# files)
CSHARP_EXTENSIONS = ['.cs', '.csx']

def find_csharp_files(root: Path) -> List[Path]:
    """Find all C# files in the workspace."""
    csharp_files = []
    for ext in CSHARP_EXTENSIONS:
        csharp_files.extend(root.rglob(f'*{ext}'))
    return sorted(csharp_files)

def remove_redundant_element_types(content: str) -> Tuple[str, List[str]]:
    """
    Remove redundant SelectedHostElementTypes code where SelectedHostCategories already exists.
    Returns: (modified_content, list_of_removals)
    """
    removals = []
    lines = content.split('\n')
    new_lines = []
    i = 0
    
    while i < len(lines):
        line = lines[i]
        
        # Check if this line has SelectedHostElementTypes
        if 'SelectedHostElementTypes' in line or 'GetSelectedHostElementTypes' in line:
            # Look ahead to see if SelectedHostCategories exists nearby (within 10 lines)
            found_categories = False
            for j in range(i, min(i + 10, len(lines))):
                if 'SelectedHostCategories' in lines[j] or 'GetSelectedHostCategories' in lines[j]:
                    found_categories = True
                    break
            
            # Also check before (within 10 lines)
            if not found_categories:
                for j in range(max(0, i - 10), i):
                    if 'SelectedHostCategories' in lines[j] or 'GetSelectedHostCategories' in lines[j]:
                        found_categories = True
                        break
            
            if found_categories:
                # This is redundant - remove it
                # Check if it's a variable declaration that spans multiple lines
                if 'var ' in line or 'List<string>' in line or '=' in line:
                    # Try to remove the entire statement (may span multiple lines)
                    statement_lines = [line]
                    j = i + 1
                    # Continue until we hit a semicolon or closing brace
                    while j < len(lines) and j < i + 5:
                        statement_lines.append(lines[j])
                        if ';' in lines[j] or '}' in lines[j]:
                            break
                        j += 1
                    
                    # Remove if it's a simple assignment or variable declaration
                    full_statement = ' '.join(statement_lines)
                    if ('GetSelectedHostElementTypes' in full_statement or 
                        'SelectedHostElementTypes' in full_statement):
                        removals.append(f"Removed redundant line {i+1}: {line.strip()[:80]}")
                        # Skip the lines we're removing
                        i = j + 1
                        continue
            
            # Not redundant, keep it for now (will be replaced later)
            new_lines.append(line)
            i += 1
        else:
            new_lines.append(line)
            i += 1
    
    return '\n'.join(new_lines), removals

def remove_redundant_patterns(content: str) -> Tuple[str, List[str]]:
    """
    Remove specific redundant patterns where both element types and categories exist.
    """
    removals = []
    original = content
    
    # Pattern 1: Remove duplicate variable declarations
    # var currentHostElementTypes = ...;
    # var currentHostCategories = ...;
    pattern1 = re.compile(
        r'var\s+currentHostElementTypes\s*=\s*[^;]+;\s*\n\s*var\s+currentHostCategories\s*=\s*[^;]+;',
        re.MULTILINE
    )
    matches = list(pattern1.finditer(content))
    for match in reversed(matches):
        # Keep only the categories line
        full_match = match.group(0)
        lines = full_match.split('\n')
        categories_line = [l for l in lines if 'currentHostCategories' in l]
        if categories_line:
            content = content[:match.start()] + categories_line[0] + '\n' + content[match.end():]
            removals.append(f"Removed redundant currentHostElementTypes declaration")
    
    # Pattern 2: Remove element types from filter object assignments when categories exist
    # selectedFilter.SelectedHostElementTypes = currentHostElementTypes;
    # selectedFilter.SelectedHostCategories = currentHostCategories;
    pattern2 = re.compile(
        r'selectedFilter\.SelectedHostElementTypes\s*=\s*[^;]+;\s*\n\s*selectedFilter\.SelectedHostCategories\s*=\s*[^;]+;',
        re.MULTILINE
    )
    matches = list(pattern2.finditer(content))
    for match in reversed(matches):
        full_match = match.group(0)
        lines = full_match.split('\n')
        categories_line = [l for l in lines if 'SelectedHostCategories' in l]
        if categories_line:
            content = content[:match.start()] + categories_line[0] + '\n' + content[match.end():]
            removals.append(f"Removed redundant SelectedHostElementTypes assignment")
    
    # Pattern 3: Remove element types from filter object property access when categories exist
    # filter?.SelectedHostElementTypes ?? ...
    # filter?.SelectedHostCategories ?? ...
    pattern3 = re.compile(
        r'filter\?\.SelectedHostElementTypes\s*[^;]+;\s*\n\s*filter\?\.SelectedHostCategories\s*[^;]+;',
        re.MULTILINE
    )
    matches = list(pattern3.finditer(content))
    for match in reversed(matches):
        full_match = match.group(0)
        lines = full_match.split('\n')
        categories_line = [l for l in lines if 'SelectedHostCategories' in l]
        if categories_line:
            content = content[:match.start()] + categories_line[0] + '\n' + content[match.end():]
            removals.append(f"Removed redundant filter?.SelectedHostElementTypes access")
    
    return content, removals

def replace_remaining_element_types(content: str) -> Tuple[str, List[str]]:
    """
    Replace remaining SelectedHostElementTypes with SelectedHostCategories.
    """
    replacements = []
    original = content
    
    # Replacement patterns
    patterns = [
        (r'SelectedHostElementTypes', 'SelectedHostCategories', 'Property name'),
        (r'selectedHostElementTypes', 'selectedHostCategories', 'Variable name'),
        (r'GetSelectedHostElementTypes', 'GetSelectedHostCategories', 'Delegate name'),
        (r'\bhostTypes\b', 'hostCategories', 'Variable: hostTypes'),
        (r'\bhostElementTypes\b', 'hostCategories', 'Variable: hostElementTypes'),
        (r'\bcurrentHostElementTypes\b', 'currentHostCategories', 'Variable: currentHostElementTypes'),
        (r'\bselectedHostTypes\b', 'selectedHostCategories', 'Variable: selectedHostTypes'),
        (r'\ballowedHostTypes\b', 'allowedHostCategories', 'Variable: allowedHostTypes'),
        (r'host element types', 'host categories', 'Comment: host element types'),
        (r'Host element types', 'Host categories', 'Comment: Host element types'),
    ]
    
    for pattern, replacement, description in patterns:
        matches = list(re.finditer(pattern, content))
        if matches:
            count = len(matches)
            content = re.sub(pattern, replacement, content)
            replacements.append(f"{description}: {count} replacement(s)")
    
    return content, replacements

def process_file(file_path: Path) -> Tuple[bool, List[str]]:
    """
    Process a single file: remove redundant code, then replace remaining occurrences.
    Returns: (was_changed, list_of_changes)
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
    except Exception as e:
        return False, [f"Error reading file: {e}"]
    
    original_content = content
    all_changes = []
    
    # Step 1: Remove redundant code
    content, removals = remove_redundant_patterns(content)
    all_changes.extend(removals)
    
    # Step 2: Remove redundant element types (line-by-line)
    content, line_removals = remove_redundant_element_types(content)
    all_changes.extend(line_removals)
    
    # Step 3: Replace remaining occurrences
    content, replacements = replace_remaining_element_types(content)
    all_changes.extend(replacements)
    
    # Only write if changes were made
    if content != original_content:
        try:
            # Create backup
            backup_path = file_path.with_suffix(file_path.suffix + '.bak')
            with open(backup_path, 'w', encoding='utf-8') as f:
                f.write(original_content)
            
            # Write updated content
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(content)
            
            return True, all_changes
        except Exception as e:
            return False, [f"Error writing file: {e}"]
    
    return False, []

def main():
    """Main execution function."""
    print("=" * 80)
    print("Replacing SelectedHostElementTypes with SelectedHostCategories")
    print("Step 1: Remove redundant code where SelectedHostCategories already exists")
    print("Step 2: Replace remaining SelectedHostElementTypes with SelectedHostCategories")
    print("=" * 80)
    print(f"Workspace: {WORKSPACE_ROOT}")
    print()
    
    # Find all C# files
    csharp_files = find_csharp_files(WORKSPACE_ROOT)
    print(f"Found {len(csharp_files)} C# files")
    print()
    
    # Filter files that contain SelectedHostElementTypes
    files_to_process = []
    for file_path in csharp_files:
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
                if 'SelectedHostElementTypes' in content or 'GetSelectedHostElementTypes' in content:
                    files_to_process.append(file_path)
        except:
            continue
    
    print(f"Found {len(files_to_process)} files containing SelectedHostElementTypes")
    print()
    
    if not files_to_process:
        print("No files to process.")
        return
    
    # Show files that will be processed
    print("Files to process:")
    for i, file_path in enumerate(files_to_process, 1):
        rel_path = file_path.relative_to(WORKSPACE_ROOT)
        print(f"  {i}. {rel_path}")
    print()
    
    # Ask for confirmation (or use --yes flag)
    import sys
    auto_yes = '--yes' in sys.argv or '-y' in sys.argv
    if not auto_yes:
        response = input("Proceed with replacements? (yes/no): ").strip().lower()
        if response not in ['yes', 'y']:
            print("Cancelled.")
            return
    else:
        print("Auto-proceeding (--yes flag provided)")
    
    print()
    print("Processing files...")
    print("-" * 80)
    
    total_files_changed = 0
    
    for file_path in files_to_process:
        rel_path = file_path.relative_to(WORKSPACE_ROOT)
        was_changed, changes = process_file(file_path)
        
        if was_changed:
            total_files_changed += 1
            print(f"\n✅ {rel_path}")
            for change in changes:
                print(f"   {change}")
        else:
            print(f"⏭️  {rel_path} (no changes needed)")
    
    print()
    print("-" * 80)
    print(f"Summary:")
    print(f"  Files processed: {len(files_to_process)}")
    print(f"  Files changed: {total_files_changed}")
    print()
    print("✅ Done! Backup files (.bak) have been created.")
    print("   Review the changes and delete .bak files if everything looks good.")

if __name__ == '__main__':
    main()
