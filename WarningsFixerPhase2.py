#!/usr/bin/env python3
"""
Phase 2: Smarter Warnings Fixer with False Positive Detection
Target: UniversalClusterService.cs (120+ warnings)
Strategy: Only fix HIGH-CONFIDENCE patterns
"""

import re
import sys
import argparse
from pathlib import Path


class Phase2Fixer:
    def __init__(self, project_root):
        self.project_root = Path(project_root)
        self.stats = {"fixed": 0, "skipped": 0, "errors": 0}
        self.changes = []
    
    def log(self, msg, level="INFO"):
        print(f"[{level}] {msg}")
    
    def read_file(self, filepath):
        try:
            return filepath.read_text(encoding='utf-8')
        except Exception as e:
            self.log(f"Failed to read {filepath}: {e}", "ERROR")
            self.stats["errors"] += 1
            return None
    
    def write_file(self, filepath, content):
        try:
            backup = filepath.with_suffix(filepath.suffix + '.backup')
            backup.write_text(filepath.read_text(encoding='utf-8'), encoding='utf-8')
            filepath.write_text(content, encoding='utf-8')
            return True
        except Exception as e:
            self.log(f"Failed to write {filepath}: {e}", "ERROR")
            self.stats["errors"] += 1
            return False
    
    def fix_cs8629_safe(self, lines):
        """CS8629: Nullable value types - Add ?.Value ?? default"""
        modified = False
        for i, line in enumerate(lines):
            # Skip if already fixed
            if '?.' in line or '??' in line:
                continue
            
            # Pattern: double x = nullableDouble;
            if re.search(r'^\s*(double|int|float|long)\s+\w+\s*=\s*\w+\s*;', line):
                match = re.match(r'^(\s*)(double|int|float|long)\s+(\w+)\s*=\s*(\w+)\s*;', line)
                if match:
                    indent, typ, var, src = match.groups()
                    # Recognize common param naming patterns too (widthParam, heightParam)
                    if src.isupper() or 'nullable' in src.lower() or src.lower().endswith('param') or src.lower().endswith('paramvalue'):
                        default = "0.0" if typ in ("double", "float") else "0"
                        new_line = f"{indent}{typ} {var} = {src}?.Value ?? {default};"
                        self.log(f"  Line {i+1}: {line.strip()[:50]}...")
                        lines[i] = new_line
                        self.changes.append((i+1, "CS8629", line.strip(), new_line.strip()))
                        self.stats["fixed"] += 1
                        modified = True
        
        return lines, modified
    
    def fix_cs8600_safe(self, lines):
        """CS8600: Null assignment - Add ?? operator"""
        modified = False
        for i, line in enumerate(lines):
            # Skip if already has null coalescing
            if '??' in line or '?.' in line:
                continue
            
            # Pattern: string x = GetValue();
            if re.search(r'string\s+\w+\s*=\s*\w+\.\w+\(\)', line) and '??' not in line:
                match = re.match(r'^(\s*)string\s+(\w+)\s*=\s*(.+)\(\)\s*;', line)
                if match:
                    indent, var, method = match.groups()
                    # Only fix if method name suggests nullable (Get prefix)
                    # If method returns a parameter-like object, or endswith GetString, etc.
                    if method.strip().endswith('.Get') or method.strip().endswith('.Get') or 'Get' in method:
                        new_line = f"{indent}string {var} = {method}() ?? string.Empty;"
                        self.log(f"  Line {i+1}: {line.strip()[:50]}...")
                        lines[i] = new_line
                        self.changes.append((i+1, "CS8600", line.strip(), new_line.strip()))
                        self.stats["fixed"] += 1
                        modified = True
        
        return lines, modified

    def fix_AsDouble_and_AsInteger(self, lines):
        """Fix common Revit API patterns where AsDouble/AsInteger/AsString are called on nullable parameter objects.
        Examples:
        double setWidth = widthParam.AsDouble();
        => double setWidth = widthParam?.AsDouble() ?? 0.0;
        string s = idParam.AsString();
        => string s = idParam?.AsString() ?? string.Empty;
        """
        modified = False
        for i, line in enumerate(lines):
            # Skip lines that already use ?. or ??
            if '?.' in line or '??' in line:
                continue

            # Double patterns
            m = re.search(r'^(\s*)(double|float|decimal)\s+(\w+)\s*=\s*(\w+)\.AsDouble\(\)\s*;', line)
            if m:
                indent, typ, var, src = m.groups()
                # Check previous 3 lines for null-checks
                context_prev = ''.join(lines[max(0, i-3):i])
                if re.search(fr'\b{src}\s*!=\s*null|\b{src}\s*==\s*null', context_prev):
                    # Skip if the code already guards
                    continue
                default = '0.0' if typ == 'double' or typ == 'float' else '0'
                new_line = f"{indent}{typ} {var} = {src}?.AsDouble() ?? {default};"
                self.log(f"  Line {i+1}: {line.strip()[:50]}...")
                lines[i] = new_line
                self.changes.append((i+1, "CS8629_AsDouble", line.strip(), new_line.strip()))
                self.stats['fixed'] += 1
                modified = True
                continue

            # Integer patterns
            m = re.search(r'^(\s*)(int|long)\s+(\w+)\s*=\s*(\w+)\.AsInteger\(\)\s*;', line)
            if m:
                indent, typ, var, src = m.groups()
                context_prev = ''.join(lines[max(0, i-3):i])
                if re.search(fr'\b{src}\s*!=\s*null|\b{src}\s*==\s*null', context_prev):
                    continue
                default = '0' if typ == 'int' else '0L'
                new_line = f"{indent}{typ} {var} = {src}?.AsInteger() ?? {default};"
                self.log(f"  Line {i+1}: {line.strip()[:50]}...")
                lines[i] = new_line
                self.changes.append((i+1, "CS8629_AsInteger", line.strip(), new_line.strip()))
                self.stats['fixed'] += 1
                modified = True
                continue

            # String patterns - AsString
            m = re.search(r'^(\s*)string\s+(\w+)\s*=\s*(\w+)\.AsString\(\)\s*;', line)
            if m:
                indent, var, src = m.groups()
                context_prev = ''.join(lines[max(0, i-3):i])
                if re.search(fr'\b{src}\s*!=\s*null|\b{src}\s*==\s*null', context_prev):
                    continue
                new_line = f"{indent}string {var} = {src}?.AsString() ?? string.Empty;"
                self.log(f"  Line {i+1}: {line.strip()[:50]}...")
                lines[i] = new_line
                self.changes.append((i+1, "CS8600_AsString", line.strip(), new_line.strip()))
                self.stats['fixed'] += 1
                modified = True
                continue

        return lines, modified
    
    def process_file(self, filepath):
        self.log(f"Processing: {filepath.name}")
        
        content = self.read_file(filepath)
        if not content:
            return False
        
        lines = content.split('\n')
        original_lines = lines.copy()
        
        # Apply safe fixes
        lines, m1 = self.fix_cs8629_safe(lines)
        lines, m2 = self.fix_cs8600_safe(lines)
        lines, m3 = self.fix_AsDouble_and_AsInteger(lines)
        
        modified = m1 or m2 or m3
        
        if modified:
            new_content = '\n'.join(lines)
            # If not in apply mode then do not write - dry run
            if not self.apply_changes:
                self.log(f"  [DRY RUN] {len(self.changes)} changes ready for {filepath.name}")
                return True
            if self.write_file(filepath, new_content):
                self.log(f"  Saved {filepath.name}")
                return True
        else:
            self.stats["skipped"] += 1
        
        return False
    
    def run(self, apply_changes: bool = False):
        print("=" * 80)
        print("Phase 2: Smarter Warnings Fixer - Safe Pattern Detection")
        print("=" * 80)
        print()
        
        target_files = [
            "Services/UniversalClusterService.cs",
            "Services/ClashZoneService.cs",
        ]
        
        self.apply_changes = apply_changes
        for rel_path in target_files:
            filepath = self.project_root / rel_path
            if filepath.exists():
                self.process_file(filepath)
            else:
                self.log(f"File not found: {rel_path}", "SKIP")
        
        # Summary
        print()
        print("=" * 80)
        print("SUMMARY")
        print("=" * 80)
        print(f"Fixed:   {self.stats['fixed']}")
        print(f"Skipped: {self.stats['skipped']}")
        print(f"Errors:  {self.stats['errors']}")
        print()
        
        if self.changes:
            print("Changes Made:")
            for line_num, warning_type, before, after in self.changes[:10]:
                print(f"  Line {line_num} ({warning_type})")
                print(f"    - {before[:70]}")
        
        print()
        print("NEXT: Run 'dotnet build' to verify changes")
        
        return self.stats["errors"] == 0


def main():
    parser = argparse.ArgumentParser(description='Phase 2 warnings fixer')
    parser.add_argument('--apply', action='store_true', help='Apply changes (writes files). Default is dry-run')
    args = parser.parse_args()

    fixer = Phase2Fixer(Path(__file__).parent)
    success = fixer.run(apply_changes=args.apply)
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
