#!/usr/bin/env python3
"""
Test Fixer - Single File Test Version
Tests warning fixes on one file with minimal changes
"""

import os
import re
import sys
from pathlib import Path
from typing import List, Tuple, Optional

class NullReferenceFixer:
    def __init__(self, project_root: str):
        self.project_root = Path(project_root)
        # TEST: Only fix one warning in ClashZoneService.cs
        self.files_to_fix = {
            "ClashZoneService.cs": {
                "CS8602": [2555],  # Just test line 2555
            }
        }
        self.stats = {"total": 0, "fixed": 0, "skipped": 0}

    def read_file(self, file_path: str) -> List[str]:
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                return f.readlines()
        except Exception as e:
            print("[ERROR] Reading " + file_path + ": " + str(e))
            return []

    def write_file(self, file_path: str, lines: List[str]) -> bool:
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.writelines(lines)
            return True
        except Exception as e:
            print("[ERROR] Writing " + file_path + ": " + str(e))
            return False

    def fix_cs8602_dereference(self, line: str) -> Optional[str]:
        """Fix dereference of possibly null reference"""
        original = line
        
        # Category.Name pattern
        if '.Category.Name' in line and '?.Category' not in line:
            line = line.replace('.Category.Name', '?.Category?.Name ?? "Unknown"')
        
        # Location.Position pattern
        elif '.Location.Position' in line and '?.Location' not in line:
            line = line.replace('.Location.Position', '?.Location?.Position')
        
        # Generic chained property (obj.prop.value)
        elif re.search(r'(\w+)\.(\w+)\.' + r'(\w+)', line) and '?.' not in line:
            match = re.search(r'(\w+)\.(\w+)\.(\w+)', line)
            if match and '?.' not in match.group(0):
                obj, prop1, prop2 = match.groups()
                old_pattern = obj + '.' + prop1 + '.' + prop2
                new_pattern = obj + '?.' + prop1 + '?.' + prop2
                line = line.replace(old_pattern, new_pattern)
        
        return line if line != original else None

    def process_file(self, file_name: str, warnings_dict: dict) -> Tuple[int, int]:
        """Process single file"""
        paths_to_try = [
            self.project_root / "Services" / file_name,
            self.project_root / "Views" / file_name,
            self.project_root / file_name
        ]
        
        file_path = None
        for p in paths_to_try:
            if p.exists():
                file_path = p
                break
        
        if not file_path:
            print("[SKIP] File not found: " + file_name)
            return 0, 0
        
        print("[FILE] Processing: " + file_name)
        lines = self.read_file(str(file_path))
        if not lines:
            return 0, 0
        
        fixed_count = 0
        skipped_count = 0
        modified_lines = lines.copy()
        
        for warning_code, line_numbers in warnings_dict.items():
            for line_num in line_numbers:
                idx = line_num - 1
                if idx >= len(lines):
                    print("[SKIP] Line " + str(line_num) + " exceeds file length")
                    continue
                
                line = lines[idx]
                print("[CONTEXT] Line " + str(line_num) + ": " + line.strip()[:70])
                
                fixed_line = None
                
                if warning_code == "CS8602":
                    fixed_line = self.fix_cs8602_dereference(line)
                
                if fixed_line and fixed_line != line:
                    modified_lines[idx] = fixed_line
                    print("   [FIXED] Applied fix")
                    print("      Before: " + line.strip()[:70])
                    print("      After:  " + fixed_line.strip()[:70])
                    fixed_count += 1
                else:
                    print("   [SKIPPED] No pattern match or already fixed")
                    skipped_count += 1
                
                print("")
        
        # DO NOT WRITE - just show what would be changed
        print("[TEST] Not saving changes (test mode)")
        print("[TEST] To apply changes, run: python WarningsFixerSimple.py")
        
        self.stats["fixed"] += fixed_count
        self.stats["skipped"] += skipped_count
        self.stats["total"] += len([v for vals in warnings_dict.values() for v in vals])
        
        return fixed_count, skipped_count

    def run(self) -> None:
        """Run fixer on test file"""
        print("=" * 80)
        print("C# Null Reference Warnings Fixer - TEST MODE (Single File)")
        print("=" * 80)
        print("[INFO] Project Root: " + str(self.project_root))
        print("[INFO] Target Files: 1 (TEST)")
        print("[INFO] Mode: READ ONLY - No changes will be saved")
        print("")
        
        total_fixed = 0
        total_skipped = 0
        
        for file_name, warnings in self.files_to_fix.items():
            fixed, skipped = self.process_file(file_name, warnings)
            total_fixed += fixed
            total_skipped += skipped
            print("")
        
        print("=" * 80)
        print("TEST SUMMARY")
        print("=" * 80)
        print("[STATS] Total Fixed (would be): " + str(total_fixed))
        print("[STATS] Total Skipped: " + str(total_skipped))
        print("[STATS] Total Targeted: " + str(self.stats["total"]))
        print("")
        print("[NEXT] If results look good, run: python WarningsFixerSimple.py")
        print("[NEXT] This will apply fixes to all targeted files")
        print("=" * 80)

def main():
    project_root = r"c:\JSE_CSharp_Projects\JSE_MEPOPENING_23"
    
    if len(sys.argv) > 1:
        project_root = sys.argv[1]
    
    fixer = NullReferenceFixer(project_root)
    fixer.run()

if __name__ == "__main__":
    main()
