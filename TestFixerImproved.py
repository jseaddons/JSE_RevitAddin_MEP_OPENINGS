#!/usr/bin/env python3
"""
Improved Test Fixer - Safer Pattern Matching
Tests warning fixes with better context awareness
"""

import os
import re
import sys
from pathlib import Path
from typing import List, Tuple, Optional

class ImprovedNullReferenceFixer:
    def __init__(self, project_root: str):
        self.project_root = Path(project_root)
        # TEST: Test a few different warning types
        self.files_to_fix = {
            "ClashZoneService.cs": {
                "CS8602": [2806],  # Test a different line - string reference
                "CS8600": [2644],  # Test a null-to-nonnull conversion
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
        """Fix dereference of possibly null reference - SAFER VERSION"""
        original = line
        
        # Skip if already has null-conditional operator
        if '?.' in line:
            return None
        
        # Skip if inside an 'as' cast
        if ' as Autodesk' in line or ' as System' in line:
            return None
        
        # Category.Name pattern
        if '.Category.Name' in line:
            return line.replace('.Category.Name', '?.Category?.Name ?? "Unknown"')
        
        # Location.Position pattern
        elif '.Location.Position' in line:
            return line.replace('.Location.Position', '?.Location?.Position')
        
        # Definition.Name pattern (common in parameters)
        elif '.Definition.Name' in line:
            return line.replace('.Definition.Name', '?.Definition?.Name ?? "Unknown"')
        
        # StorageType pattern
        elif '.StorageType' in line and 'ParameterType' not in line:
            return line.replace('.StorageType', '?.StorageType')
        
        return None

    def fix_cs8600_null_to_nonnull(self, line: str) -> Optional[str]:
        """Fix converting null to non-nullable - SAFER VERSION"""
        # Only fix if it's a simple variable assignment
        if re.search(r'string\s+\w+\s*=\s*\w+\s*\?\..*\s*;', line):
            if '??' not in line:
                return re.sub(r'(\s*;)$', r' ?? string.Empty\1', line)
        
        return None

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
        
        for warning_code, line_numbers in warnings_dict.items():
            for line_num in line_numbers:
                idx = line_num - 1
                if idx >= len(lines):
                    print("[SKIP] Line " + str(line_num) + " exceeds file length")
                    skipped_count += 1
                    continue
                
                line = lines[idx]
                print("")
                print("[TEST] Line " + str(line_num) + " (" + warning_code + "):")
                print("      Code: " + line.rstrip()[:80])
                
                fixed_line = None
                
                if warning_code == "CS8602":
                    fixed_line = self.fix_cs8602_dereference(line)
                elif warning_code == "CS8600":
                    fixed_line = self.fix_cs8600_null_to_nonnull(line)
                
                if fixed_line and fixed_line != line:
                    print("   [WOULD FIX]")
                    print("      Before: " + line.rstrip()[:80])
                    print("      After:  " + fixed_line.rstrip()[:80])
                    fixed_count += 1
                else:
                    print("   [NO MATCH] - Pattern doesn't apply or already fixed")
                    skipped_count += 1
        
        self.stats["fixed"] += fixed_count
        self.stats["skipped"] += skipped_count
        self.stats["total"] += len([v for vals in warnings_dict.values() for v in vals])
        
        return fixed_count, skipped_count

    def run(self) -> None:
        """Run improved fixer on test file"""
        print("=" * 80)
        print("C# Null Reference Warnings Fixer - IMPROVED TEST MODE")
        print("=" * 80)
        print("[INFO] Project Root: " + str(self.project_root))
        print("[INFO] Mode: TEST (No changes saved)")
        print("[INFO] Goal: Verify pattern matching safety and accuracy")
        print("")
        
        total_fixed = 0
        total_skipped = 0
        
        for file_name, warnings in self.files_to_fix.items():
            fixed, skipped = self.process_file(file_name, warnings)
            total_fixed += fixed
            total_skipped += skipped
        
        print("")
        print("=" * 80)
        print("IMPROVED TEST RESULTS")
        print("=" * 80)
        print("[STATS] Patterns That Would Fix: " + str(total_fixed))
        print("[STATS] Patterns That Don't Apply: " + str(total_skipped))
        print("[STATS] Total Tested: " + str(self.stats["total"]))
        print("")
        print("[ASSESSMENT]")
        if total_skipped == 0:
            print("  All tested patterns matched successfully!")
        elif total_fixed == 0:
            print("  No patterns matched - may need to adjust detection")
        else:
            success_rate = (total_fixed / (total_fixed + total_skipped) * 100)
            print("  Success Rate: " + str(round(success_rate, 1)) + "%")
        print("")
        print("[NEXT STEPS]")
        print("  1. If results look good: python WarningsFixerSimple.py")
        print("  2. Then verify: dotnet build JSE_RevitAddin_MEP_OPENINGS.csproj")
        print("  3. Check warnings: Verify warning count decreased")
        print("=" * 80)

def main():
    project_root = r"c:\JSE_CSharp_Projects\JSE_MEPOPENING_23"
    
    if len(sys.argv) > 1:
        project_root = sys.argv[1]
    
    fixer = ImprovedNullReferenceFixer(project_root)
    fixer.run()

if __name__ == "__main__":
    main()
