#!/usr/bin/env python3
"""
Wrap all File.AppendAllText/WriteAllText calls with DeploymentConfiguration.DeploymentMode checks.
Handles nested blocks, different indentation levels, and complex code structures.
"""

import re
import os
from pathlib import Path

def is_protected_by_deployment_check(lines, line_index):
    """
    Check if a File.AppendAllText/WriteAllText call is already protected by a deployment mode check.
    Looks backwards up to 20 lines for the check, handling nested blocks.
    """
    if line_index < 0:
        return False
    
    # Look backwards for deployment mode check
    search_start = max(0, line_index - 20)
    found_if_statement = False
    found_deployment_check = False
    brace_count = 0
    
    for i in range(line_index, search_start - 1, -1):
        line = lines[i].strip()
        
        # Skip empty lines and comments
        if not line or line.startswith('//') or line.startswith('/*') or line.startswith('*'):
            continue
            
        # Check if this is our target line
        if i == line_index:
            # Count opening braces before this line
            for j in range(search_start, i):
                brace_count += lines[j].count('{') - lines[j].count('}')
            continue
        
        # Track braces
        brace_count += line.count('{') - line.count('}')
        
        # Look for deployment mode check pattern
        if re.search(r'if\s*\(\s*!?\s*DeploymentConfiguration\.DeploymentMode', line):
            # Found the check - verify it's for file operations
            # Check if there's a File.AppendAllText/WriteAllText after this if
            for j in range(i + 1, min(i + 10, line_index + 1)):
                if j >= len(lines):
                    break
                if re.search(r'File\.(AppendAllText|WriteAllText)', lines[j]):
                    # Check if brace balance allows this to protect our target line
                    temp_brace = 0
                    for k in range(i, j + 1):
                        temp_brace += lines[k].count('{') - lines[k].count('}')
                    if temp_brace >= 0:  # We're within the if block
                        return True
            found_deployment_check = True
        
        # If we've gone too far back (more closing braces than opening), we're out of scope
        if brace_count < 0:
            break
    
    return False

def get_indentation(line):
    """Get the indentation level (spaces/tabs) of a line."""
    return len(line) - len(line.lstrip())

def wrap_file_logging_call(lines, line_index):
    """
    Wrap a File.AppendAllText/WriteAllText call with deployment mode check.
    Returns modified line content and whether wrapping was successful.
    """
    original_line = lines[line_index]
    indent = get_indentation(original_line)
    indent_str = ' ' * indent
    
    # Extract just the content (without trailing newline)
    line_content = original_line.rstrip('\n\r')
    stripped = line_content.strip()
    
    # Check if line starts with "try {"
    starts_with_try = re.match(r'^\s*try\s*\{', stripped)
    
    # Check if it's already inside a try block (try { is on a previous line)
    in_try_block = False
    try_indent = indent
    for i in range(max(0, line_index - 15), line_index):
        if re.search(r'\btry\s*\{', lines[i]):
            in_try_block = True
            try_indent = get_indentation(lines[i])
            break
    
    # Check if this line is inside an existing if (!DeploymentConfiguration.DeploymentMode) block
    # Look backwards for the if statement
    for i in range(line_index - 1, max(0, line_index - 15), -1):
        if re.search(r'if\s*\(\s*!?\s*DeploymentConfiguration\.DeploymentMode', lines[i]):
            # Already protected, don't wrap again
            return [original_line], False
    
    # Pattern 1: Line that starts with "try { File.AppendAllText(...) } catch"
    if starts_with_try:
        # Simple pattern: try { File.AppendAllText(...); } catch { }
        # We'll wrap the File call inside the try block
        try_match = re.match(r'^\s*try\s*\{\s*(.*?)\s*\}\s*catch', stripped, re.DOTALL)
        if try_match:
            content = try_match.group(1)
            # Extract File call from content
            file_match = re.search(r'File\.(AppendAllText|WriteAllText)\([^)]+\);?', content)
            if file_match:
                before_try = stripped[:stripped.find('try')]
                catch_part = stripped[stripped.find('} catch'):]
                file_call = file_match.group(0).rstrip(';')
                
                wrapped_lines = [
                    indent_str + before_try.strip() + 'try {\n',
                    indent_str + '    // DEPLOYMENT MODE: Skip file writes\n',
                    indent_str + '    if (!DeploymentConfiguration.DeploymentMode)\n',
                    indent_str + '    {\n',
                    indent_str + '        ' + file_call + ';\n',
                    indent_str + '    }\n',
                    indent_str + catch_part + '\n'
                ]
                return wrapped_lines, True
    
    # Pattern 2: Line inside try block - wrap just the File call
    if in_try_block:
        # Preserve original indentation, add deployment check
        # The line is already indented properly inside try block
        # We'll add deployment check with same indent level
        wrapped_lines = [
            indent_str + '// DEPLOYMENT MODE: Skip file writes\n',
            indent_str + 'if (!DeploymentConfiguration.DeploymentMode)\n',
            indent_str + '{\n',
            indent_str + '    ' + stripped + '\n',
            indent_str + '}\n'
        ]
        return wrapped_lines, True
    
    # Pattern 3: Standalone line (not in try-catch)
    # For System.IO.File.AppendAllText, just use File.AppendAllText pattern
    if 'System.IO.' in stripped:
        stripped = stripped.replace('System.IO.', '')
    
    # Simple wrap
    wrapped_lines = [
        indent_str + '// DEPLOYMENT MODE: Skip file writes\n',
        indent_str + 'if (!DeploymentConfiguration.DeploymentMode)\n',
        indent_str + '{\n',
        indent_str + '    ' + stripped + '\n',
        indent_str + '}\n'
    ]
    return wrapped_lines, True

def process_file(file_path):
    """Process a single C# file and wrap unwrapped file logging calls."""
    try:
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            lines = f.readlines()
    except Exception as e:
        print(f"Error reading {file_path}: {e}")
        return False
    
    modified = False
    new_lines = []
    i = 0
    
    while i < len(lines):
        line = lines[i]
        
        # Check if this line contains File.AppendAllText or File.WriteAllText
        # Handle multi-line calls (check current line and next few lines)
        has_file_call = re.search(r'File\.(AppendAllText|WriteAllText)', line)
        multi_line_call = False
        if has_file_call:
            # Check if call continues on next lines (unclosed parenthesis)
            paren_open = line.count('(') - line.count(')')
            if paren_open > 0:
                # May be multi-line, check next lines
                for j in range(i + 1, min(i + 5, len(lines))):
                    paren_open += lines[j].count('(') - lines[j].count(')')
                    if paren_open <= 0:
                        break
                multi_line_call = paren_open > 0
            # Check if already protected
            if not is_protected_by_deployment_check(lines, i):
                preview = line.strip()
                if len(preview) > 80:
                    preview = preview[:80] + "..."
                print(f"  Line {i+1}: Found unwrapped call: {preview}")
                
                # For multi-line calls, we'll process the whole block
                if multi_line_call:
                    # Find where the call ends
                    paren_open = line.count('(') - line.count(')')
                    call_end_line = i
                    for j in range(i + 1, min(i + 10, len(lines))):
                        paren_open += lines[j].count('(') - lines[j].count(')')
                        call_end_line = j
                        if paren_open <= 0:
                            break
                    # Process the first line, skip the rest
                    wrapped, success = wrap_file_logging_call(lines, i)
                    if success:
                        new_lines.extend(wrapped)
                        # Skip all lines in the multi-line call
                        i = call_end_line + 1
                        modified = True
                        continue
                
                # Try to wrap it
                wrapped, success = wrap_file_logging_call(lines, i)
                if success:
                    # Add wrapped lines (they already have newlines)
                    new_lines.extend(wrapped)
                    modified = True
                    # Continue to next line (don't add original line)
                    i += 1
                    continue
                else:
                    # Couldn't wrap, keep original
                    print(f"    WARNING: Could not wrap, keeping original")
                    new_lines.append(line)
            else:
                # Already protected, keep as is
                pass  # Don't print, already wrapped
                new_lines.append(line)
        else:
            # Not a file logging call, keep as is
            new_lines.append(line)
        
        i += 1
    
    if modified:
        # Backup original file
        backup_path = str(file_path) + '.backup'
        try:
            with open(backup_path, 'w', encoding='utf-8') as f:
                f.writelines(lines)
            print(f"  OK: Created backup: {backup_path}")
        except:
            pass
        
        # Write modified file
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.writelines(new_lines)
            print(f"  OK: Modified and saved: {file_path}")
            return True
        except Exception as e:
            print(f"  ERROR: Error writing {file_path}: {e}")
            return False
    
    return False

def main():
    """Main function to process all C# files in Services directory."""
    import sys
    
    # TEST MODE: Process only one file
    test_file = Path('Services/UniversalSleevePlacerService.cs')
    
    if len(sys.argv) > 1:
        test_file = Path(sys.argv[1])
    
    if not test_file.exists():
        print(f"Error: Test file not found at {test_file.absolute()}")
        return
    
    print(f"TEST MODE: Processing single file: {test_file}\n")
    
    if process_file(test_file):
        print(f"\n{'='*60}")
        print(f"OK: File modified successfully!")
        print(f"Please build the project to verify no compilation errors.")
        print(f"{'='*60}")
    else:
        print(f"\n{'='*60}")
        print(f"INFO: No changes needed or file processing failed.")
        print(f"{'='*60}")

if __name__ == '__main__':
    main()

